import csv
from itertools import chain, count
from pathlib import Path
from contextlib import ExitStack
import openpyxl as xl
import re


def strip_csv_whitespace(csv_filename: Path) -> None:
    """Strips leading and trailing whitespace from CSV values.

    Operates in-place on a file.

    Args:
        csv_filename: Path to CSV file to strip.
    """

    cleaned_rows = []

    with open(csv_filename, 'r') as csv_file_obj_r:
        for row in csv.reader(csv_file_obj_r):
            cleaned_row = []
            for value in row:
                try:
                    cleaned_row.append(value.strip())
                except AttributeError:
                    cleaned_row.append(value)
            cleaned_rows.append(cleaned_row)

    with open(csv_filename, 'w') as csv_file_obj_w:
        writer = csv.writer(csv_file_obj_w)
        writer.writerows(cleaned_rows)


def sanitize_csv_column_names(csv_filename: Path) -> None:
    """Sanitize CSV column names into ones suitable to be SQL identifiers.

    Args:
        csv_filename: Path of CSV to sanitize columns within.

    Raises:
        ValueError: An empty column name was encountered.
        ValueError: A column name that started with a number was encountered.
    """
    with open(csv_filename, 'r') as csv_file_obj_r:
        reader = csv.reader(csv_file_obj_r)
        headers = next(reader)
        rows = list(reader)

    sanitized_headers = []
    for header in headers:
        h = str(header).strip().lower()
        if not h:
            raise ValueError(
                f'Encountered an empty column name in {csv_filename}. Please fix this manually.'
            )
        if h[0].isnumeric():
            raise ValueError(
                'Encountered column name that starts with a number. Please fix this manually.'
            )
        h = ''.join(c if c.isalnum() else '_' for c in h)
        sanitized_headers.append(h)

    with open(csv_filename, 'w') as csv_file_obj_w:
        writer = csv.writer(csv_file_obj_w)
        writer.writerow(sanitized_headers)
        writer.writerows(rows)


def convert_excel_file_to_csvs(
    xlsx_filename: Path,
    strip_whitespace: bool = True,
    sanitize_headers: bool = True,
) -> None:
    """Convert an Excel file to CSV format.

    Each independent worksheet within an Excel workbook will become its own CSV.

    Args:
        filename: Path to Excel file to convert.
    """

    wb = xl.load_workbook(xlsx_filename, read_only=True)
    for worksheet in wb.worksheets:
        csv_name = xlsx_filename.parent / (
            xlsx_filename.stem.replace(' ', '_') + f'_{worksheet.title}.csv'
        )
        with open(csv_name, 'w') as csv_file_obj:
            writer = csv.writer(csv_file_obj)
            for row in worksheet.values:
                writer.writerow(row)

        if strip_whitespace:
            strip_csv_whitespace(csv_name)
        if sanitize_headers:
            sanitize_csv_column_names(csv_name)



def merge_all_csv_in_dir(
        input_dir: Path,
        output_dir: Path = Path('./output/'),
        append_page_line: bool = False,
        filename_regex: re.Pattern | None = None
    ):
    """Aggregate all CSV files within a directory into a new CSV.

    Args:
        input_dir: Directory holding CSV files to be aggregated.
        output_dir: Output directory to hold the aggregated CSV file. Defaults to Path('./output/').
        append_page_line: Append page and line numbers to the aggregated CSV file. These can be used to couple targets and responses.
                          Only use this option if the page and line numbers are not already provided in the CSV files.
        filename_regex: Regex to match filenames by. This must be a regex that provides two named groups for the page and line, named as such.

    Raises:
        FileNotFoundError: The input directory did not exist.
        ValueError: The input directory was not a directory.
        OSError: An issue occurred when making a previously nonexisting output directory.
    """
    if not input_dir.exists():
        raise FileNotFoundError('Input directory does not exist')
    if not input_dir.is_dir():
        raise ValueError('Input directory is not actually a directory')
    if append_page_line and filename_regex is None:
        raise ValueError('Cannot append page and line information without filename regex.')

    try:
        output_dir.mkdir(exist_ok=True, parents=True)
    except OSError as e:
        raise OSError('Issue occurred making output directory') from e

    csv_filenames = list(input_dir.glob('*.csv'))
    merged_filename = csv_filenames[0].name.split('_')[0] + '_merged.csv'

    with ExitStack() as stack:
        csv_files = {csv_file: stack.enter_context(open(csv_file, 'r')) for csv_file in csv_filenames}
        output_file = stack.enter_context(open(output_dir / merged_filename, 'w'))

        readers = {name: csv.DictReader(fp) for name, fp in csv_files.items()}
        all_headers = chain(*(r.fieldnames for r in readers.values()))  # type: ignore
        seen_headers: set[str] = set()
        seen_headers_add = seen_headers.add
        headers_no_duplicates = [h for h in all_headers if not(h in seen_headers or seen_headers_add(h))]
        if append_page_line:
            headers_no_duplicates.append('page')
            headers_no_duplicates.append('line')

        writer = csv.DictWriter(output_file, fieldnames=headers_no_duplicates)
        writer.writeheader()
        for name, reader in readers.items():
            for row in reader:
                row_to_write = {k: v for k,v in row.items()}
                if append_page_line:
                    # filename_regex is guaranteed to not be None at this point
                    match = filename_regex.match(name)  # type: ignore
                    if match is None:
                        raise ValueError(f'The following filename failed to match against the provided regex: {name}')
                    row_to_write |= match.groupdict()
                try:
                    writer.writerow(row_to_write)
                except:
                    print(name)
                    print(row_to_write)
                    raise




if __name__ == '__main__':
    ...
    gullah = Path('..') / 'data' / 'Gullah' / 'Text_files'
    gullah_regex = re.compile(r'Gullah_page(?P<page>[a-z0-9]*)line(?P<line>[a-z0-9]*)_Sheet[0-9]*\.csv')

    for file in gullah.glob('*.xlsx'):
        convert_excel_file_to_csvs(file)

    merge_all_csv_in_dir(gullah)

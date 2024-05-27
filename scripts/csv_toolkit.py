import argparse
import csv
import re
from contextlib import ExitStack
from itertools import chain
from pathlib import Path

import openpyxl as xl


def obtain_choice_from_user(choices: list[str]) -> str:
    """Obtain a choice from the user via the command line.

    Args:
        choices: The choices the user can provide.

    Raises:
        ValueError: Less than 2 choices were provided.

    Returns:
        The choice the user selected. Normalized to be lowercase and stripped.
    """
    if len(choices) <= 1:
        raise ValueError('Cannot make a choice out of 1 or fewer options.')

    choices = [c.strip().lower() for c in choices]
    options_str = (
        ', '.join(f'[{c}]' for c in choices[:-1])
        + f'{',' if len(choices) > 2 else ''} or [{choices[-1]}]'
    )

    while (
        choice := input(f'Do you wish to {options_str} this column? ').strip().lower()
    ) not in choices:
        print(f'Please choose one of {options_str}.\n')

    return choice


def verify_new_column_name() -> str:
    """Verify a new column name from the user via the command line.

    Returns:
        The name column name, sanitized.
    """
    while True:
        name = (
            input('Please enter a new column name (case insensitive): ').strip().lower()
        )
        if not name.isidentifier():
            print(
                'New name is not a valid identifier. '
                'Please only use characters alphanumeric characters and underscores. '
                'The first character cannot be a number.\n'
            )
            continue
        if (
            name
            == input('Please confirm the new column name (case insensitive): ')
            .strip()
            .lower()
        ):
            return name
        print('New name and confirmation did not match.\n')


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
    """Sanitize CSV column names into ones suitable to be identifiers.
    For these purposes, a suitable identifier is a valid Python identifier.

    Operates in-place on a file.

    Args:
        csv_filename: Path of CSV to sanitize columns within.
    """
    with open(csv_filename, 'r') as csv_file_obj_r:
        reader = csv.reader(csv_file_obj_r)
        headers = next(reader)
        rows = list(reader)

    sanitized_headers: list[str] = []
    empty_column_header_indices: list[int] = []
    for column_index, header in enumerate(headers):
        h = str(header).strip().lower()
        if not h:
            empty_column_header_indices.append(column_index)
        if h and h[0].isnumeric():
            print(
                'Encountered column name that starts with a number. This is not allowed.'
            )
            h = verify_new_column_name()
        h = ''.join(c if c.isalnum() else '_' for c in h)
        sanitized_headers.append(h)

    discarded_column_indices: list[int] = []
    for column_index in empty_column_header_indices:
        print(
            f'Encountered an empty column name in {csv_filename}. A sample of the column is provided below. '
        )
        column_sample: list[str] = []
        row_index = 0
        while len(column_sample) != min(
            12, len(rows)
        ):  # TODO: make this 12 an argument?
            if (entry := rows[row_index][column_index]) != '':
                column_sample.append(entry)
            row_index += 1
        print(' | '.join(column_sample))

        choice = obtain_choice_from_user(['rename', 'discard'])
        if choice == 'rename':
            sanitized_headers[column_index] = verify_new_column_name()
        elif choice == 'discard':
            discarded_column_indices.append(column_index)
        print()

    for column_index in reversed(discarded_column_indices):
        sanitized_headers.pop(column_index)
        for row in rows:
            row.pop(column_index)

    with open(csv_filename, 'w') as csv_file_obj_w:
        writer = csv.writer(csv_file_obj_w)
        writer.writerow(sanitized_headers)
        writer.writerows(rows)


def prune_empty_csv_columns(csv_filename: Path) -> None:
    """Prune empty columns from a CSV file.
    Columns with empty header names are pruned silently. If a column name is found, the user is prompted.

    Operates in-place on a file.

    Args:
        csv_filename: Path of CSV file to prune.
    """
    with open(csv_filename, 'r') as csv_file_obj_r:
        reader = csv.reader(csv_file_obj_r)
        headers = next(reader)
        rows = list(reader)

    empty_column_indices = []
    for column_index, header in enumerate(headers):
        for row in rows:
            if row[column_index] != '':
                break
        else:
            if header == '':
                empty_column_indices.append(column_index)
            else:
                print(f'Column "{header}" was found to be empty.')
                choice = obtain_choice_from_user(['discard', 'keep'])
                if choice == 'discard':
                    empty_column_indices.append(column_index)
                print()

    new_rows = []
    for row in chain([headers], rows):
        new_row = row[:]
        for empty_index in reversed(empty_column_indices):
            new_row.pop(empty_index)
        new_rows.append(new_row)

    with open(csv_filename, 'w') as csv_file_obj_w:
        writer = csv.writer(csv_file_obj_w)
        writer.writerows(new_rows)


def append_metadata_from_filename(
    csv_filename: Path,
    filename_regex: re.Pattern,
) -> None:
    """Append metadata columns from a CSV filename.

    This appends columns with info on the project, page number, and line number to the CSV from its filename.
    If any of these columns already exist, the user will be prompted to keep or overwrite the column based on it's contents.
    Operates in-place on a file.

    Args:
        csv_filename: Path of CSV file to append metadata to.
        filename_regex: Regex to match filenames by.
                This must be a regex that provides only three named groups for the project, page and line, named as such.
                This regex will be matched against CSV files, not Excel files.
    """
    # TODO: write this.


def convert_excel_file_to_csvs(
    xlsx_filename: Path,
    strip_whitespace: bool,
    sanitize_headers: bool,
    prune_empty_columns: bool,
    append_metadata: bool,
    filename_regex: re.Pattern,
    output_dir: Path,
) -> None:
    """Convert an Excel file to CSV format.
    Each independent worksheet within an Excel workbook will become its own CSV.

    Args:
        xlsx_filename: Name of Excel file to convert.
        strip_whitespace: Option to strip whitespace.
        sanitize_headers: Option to sanitize headers.
        prune_empty_columns: Option to prune empty columns.
        append_metadata: Option to append file metadata to the CSV based on its filename. Defaults to False.
        filename_regex: Regex to match filenames by.
                        This must be a regex that provides only three named groups for the project, page and line, named as such.
                        This regex will be matched against CSV files, not Excel files. Remember to account for this if providing alternate regex.
        output_dir: Output directory to place converted files into.
                    The script will create the output directory if it does not exist.
    """

    output_dir.mkdir(parents=True, exist_ok=True)

    wb = xl.load_workbook(xlsx_filename, read_only=True)
    for worksheet in wb.worksheets:
        csv_name = output_dir / (
            xlsx_filename.stem.replace(' ', '_') + f'_{worksheet.title}.csv'
        )
        with open(csv_name, 'w') as csv_file_obj:
            writer = csv.writer(csv_file_obj)
            for row in worksheet.values:
                writer.writerow(row)

        if strip_whitespace:
            strip_csv_whitespace(csv_name)
        if prune_empty_columns:
            prune_empty_csv_columns(csv_name)
        if sanitize_headers:
            sanitize_csv_column_names(csv_name)
        if append_metadata:
            append_metadata_from_filename(csv_name, filename_regex)


def merge_all_csv_in_dir(
    input_dir: Path,
    output_dir: Path,
) -> None:
    """Aggregate all CSV files within a directory into a new CSV.

    Args:
        input_dir: Directory holding CSV files to be aggregated.
        output_dir: Output directory to hold the aggregated CSV file. Defaults to Path('./output/').

    Raises:
        FileNotFoundError: The input directory did not exist.
        ValueError: The input directory was not a directory.
        OSError: An issue occurred when making a previously nonexisting output directory.
    """
    if not input_dir.exists():
        raise FileNotFoundError('Input directory does not exist')
    if not input_dir.is_dir():
        raise ValueError('Input directory is not actually a directory')

    try:
        output_dir.mkdir(exist_ok=True, parents=True)
    except OSError as e:
        raise OSError('Issue occurred making output directory') from e

    csv_filenames = list(input_dir.glob('*.csv'))
    merged_filename = csv_filenames[0].name.split('_')[0] + '_merged.csv'

    with ExitStack() as stack:
        csv_files = {
            csv_file: stack.enter_context(open(csv_file, 'r'))
            for csv_file in csv_filenames
        }
        output_file = stack.enter_context(open(output_dir / merged_filename, 'w'))

        readers = {name: csv.DictReader(fp) for name, fp in csv_files.items()}
        all_headers = chain(*(r.fieldnames for r in readers.values()))  # type: ignore
        seen_headers: set[str] = set()
        seen_headers_add = seen_headers.add
        headers_no_duplicates = [
            h for h in all_headers if not (h in seen_headers or seen_headers_add(h))
        ]

        writer = csv.DictWriter(output_file, fieldnames=headers_no_duplicates)
        writer.writeheader()
        for name, reader in readers.items():
            for row in reader:
                row_to_write = {k: v for k, v in row.items()}
                try:
                    writer.writerow(row_to_write)
                except:
                    print(name)
                    print(row_to_write)
                    raise


def build_arg_parser() -> argparse.ArgumentParser:
    """Build command line arguments.

    Returns:
        Parser for command line arguments.
    """
    parser = argparse.ArgumentParser()

    parser.add_argument(
        '-w',
        '--strip_whitespace',
        action='store_true',
        help='strip leading and trailing whitespace from CSV entries',
    )
    parser.add_argument(
        '-s',
        '--sanitize_headers',
        action='store_true',
        help='sanitize CSV column headers',
    )
    parser.add_argument(
        '-p',
        '--prune_empty_columns',
        action='store_true',
        help='prune empty columns from CSV',
    )
    parser.add_argument(
        '-a',
        '--append_metadata',
        action='store_true',
        help='append filename metadata to CSV',
    )
    parser.add_argument(
        '-f',
        '--filename_regex',
        type=lambda s: re.compile(s),
        default=re.compile(
            r'(?P<project>[a-zA-Z]*)'
            r'_page(?P<page>[a-z0-9]*)'
            r'line(?P<line>[a-z0-9]*)'
            r'_Sheet[0-9]*\.csv'
        ),
        help=(
            'expected regex CSV filenames fit under. Must provide named groups for "project", "line", and "page" '
            '(default: (?P<project>[a-zA-Z]*)_page(?P<page>[a-z0-9]*)line(?P<line>[a-z0-9]*)_Sheet[0-9]*\\.csv )'
        ),
    )
    parser.add_argument(
        '-m',
        '--merge',
        action='store_true',
        help='merge all processed CSV files into one aggregate file',
    )
    parser.add_argument(
        '-o',
        '--output_directory',
        type=Path,
        default=Path('./output/'),
        help='output directory for processed Excel files (default: ./output/ )',
    )
    parser.add_argument(
        'input_directory',
        type=Path,
        help='directory of Excel (xlsx) files to process',
    )

    return parser


def main() -> None:
    parser = build_arg_parser()
    args = parser.parse_args()
    # print(args)
    # return
    for file in args.input_directory.glob('*.xlsx'):
        print(file)
        convert_excel_file_to_csvs(
            file,
            strip_whitespace=args.strip_whitespace,
            sanitize_headers=args.sanitize_headers,
            prune_empty_columns=args.prune_empty_columns,
            append_metadata=args.append_metadata,
            filename_regex=args.filename_regex,
            output_dir=args.output_directory,
        )
        break

    # merge_all_csv_in_dir(gullah)


if __name__ == '__main__':
    main()

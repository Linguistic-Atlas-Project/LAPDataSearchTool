import argparse
import csv
import re
import warnings
from contextlib import ExitStack
from itertools import chain
from pathlib import Path

import openpyxl as xl


# TODO: functionality to stop script early and rerun later from same spot?

ALLOWED_COLUMN_NAMES = [
    'informant',
    'response',
    'comments',
    'phonetic_transcription',
    'project',
    'page',
    'line',
    'filename',
]
remembered_enforced_column_name_changes: dict[str, int] = {}


def obtain_choice_from_user(choices: list[str], message: str) -> str:
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

    while (choice := input(message.format(options_str)).strip().lower()) not in choices:
        print(f'Please choose one of {options_str}.\n')

    return choice


def verify_new_column_name(enforce: bool, original_column_name: str) -> str:
    """Verify a new column name from the user via the command line.

    Returns:
        The name column name, sanitized.
    """
    if not enforce:
        while True:
            name = (
                input('Please enter a new column name (case insensitive): ')
                .strip()
                .lower()
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
    else:
        if original_column_name in remembered_enforced_column_name_changes:
            return ALLOWED_COLUMN_NAMES[
                remembered_enforced_column_name_changes[original_column_name]
            ]

        print(f'A disallowed column name was encountered: "{original_column_name}".')
        print('Please choose the closest allowed column name, or discard this column.')
        print(
            'The following are allowable column names. Please make your choice by number.'
        )
        print('0. DISCARD THIS COLUMN')
        for i, c in enumerate(ALLOWED_COLUMN_NAMES):
            print(f'{i+1}. {c.replace('_', ' ').title()}')

        while True:
            try:
                number = (
                    int(input('Please input the number of your choice: ').strip()) - 1
                )
            except ValueError:
                print('That was not a number.\n')
                continue

            if number not in range(-1, len(ALLOWED_COLUMN_NAMES)):
                print('That was not a valid choice.\n')
                continue

            choice = obtain_choice_from_user(
                ['yes', 'no'], 'Remember this decision? {} '
            )
            print()
            if choice == 'yes':
                remembered_enforced_column_name_changes[original_column_name] = number

            return ALLOWED_COLUMN_NAMES[number] if number != -1 else ''


def print_column_sample(rows: list[list[str]], column_index: int) -> None:
    column_sample: list[str] = []

    row_index = 0
    num_sample_entries = min(12, len(rows))  # TODO: make this 12 an argument?

    while len(column_sample) < num_sample_entries:
        if (entry := rows[row_index][column_index]) != '':
            column_sample.append(entry)
        row_index += 1
        if row_index == len(rows):
            break

    print(' | '.join(column_sample))


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


def sanitize_csv_column_names(
    csv_filename: Path,
    destructive: bool,
    discard_empty: bool,
    assume_headers: bool,
    enforce_headers: bool,
) -> None:
    """Sanitize CSV column names into ones suitable to be identifiers.
    For these purposes, a suitable identifier is a valid Python identifier.

    Operates in-place on a file.

    Args:
        csv_filename: Path of CSV to sanitize columns within.
        destructive: Assuming everything unsanitary should be discarded.
        assume_headers: Assume the first line of CSV file is headers.
    """
    with open(csv_filename, 'r') as csv_file_obj_r:
        reader = csv.reader(csv_file_obj_r)
        headers = next(reader)
        rows = list(reader)

    discarded_column_indices: list[int] = []

    if not assume_headers:
        print('The following is the first row of the file:\n', ' | '.join(headers))
        choice = obtain_choice_from_user(
            ['yes', 'no'], message='Please confirm if these are headers {}: '
        )
        if choice == 'no':
            rows.insert(0, headers)
            new_headers = []
            print(
                'Please provide columns names for these columns, in order from left to right, one at a time.'
            )
            for column_index in range(len(headers)):
                print('A sample of the column is provided below')
                print_column_sample(rows, column_index)
                choice = obtain_choice_from_user(
                    ['rename', 'discard'], 'Do you wish to {} this column? '
                )
                if choice == 'rename':
                    new_headers.append(
                        verify_new_column_name(enforce_headers, headers[column_index])
                    )
                elif choice == 'discard':
                    discarded_column_indices.append(column_index)
                print()

            headers = new_headers

    sanitized_headers: list[str] = []
    empty_column_header_indices: list[int] = []

    for column_index, header in enumerate(headers):
        if column_index in discarded_column_indices:
            continue
        h = str(header).strip().lower()
        if not h:
            empty_column_header_indices.append(column_index)
        elif h not in ALLOWED_COLUMN_NAMES and enforce_headers:
            h = verify_new_column_name(enforce_headers, h)
        elif h[0].isnumeric():
            if not destructive:
                print(
                    f'Encountered column name that starts with a number: {h}. This is not allowed.'
                )
                choice = obtain_choice_from_user(
                    ['rename', 'discard'], 'Do you wish to {} this column? '
                )
            else:
                choice = 'discard'

            if choice == 'rename':
                h = verify_new_column_name(enforce_headers, h)
            elif choice == 'discard':
                discarded_column_indices.append(column_index)

        h = ''.join(c if c.isalnum() else '_' for c in h)
        sanitized_headers.append(h)

    for column_index in empty_column_header_indices:
        if not discard_empty:
            print(
                f'Encountered an empty column name in {csv_filename}. '
                'A sample of the column is provided below. '
            )
            print_column_sample(rows, column_index)
            choice = obtain_choice_from_user(
                ['rename', 'discard'], 'Do you wish to {} this column? '
            )
            print()
        else:
            choice = 'discard'

        if choice == 'rename':
            sanitized_headers[column_index] = verify_new_column_name(
                enforce_headers, sanitized_headers[column_index]
            )
        elif choice == 'discard':
            discarded_column_indices.append(column_index)

    for column_index in reversed(sorted(discarded_column_indices)):
        sanitized_headers.pop(column_index)
        for row in rows:
            row.pop(column_index)

    with open(csv_filename, 'w') as csv_file_obj_w:
        writer = csv.writer(csv_file_obj_w)
        writer.writerow(sanitized_headers)
        writer.writerows(rows)


def prune_empty_csv_rows(csv_filename: Path) -> None:
    """Prune empty columns from a CSV file.

    Operates in-place on a file.

    Args:
        csv_filename: Path of CSV file to prune.
    """
    with open(csv_filename, 'r') as csv_file_obj_r:
        reader = csv.reader(csv_file_obj_r)
        new_rows = [row for row in reader if not all('' == s for s in row)]

    with open(csv_filename, 'w') as csv_file_obj_w:
        writer = csv.writer(csv_file_obj_w)
        writer.writerows(new_rows)


def prune_padding_csv_columns(csv_filename: Path) -> None:
    """Prune empty columns that have no header from a CSV file.

     Operates in-place on a file.

    Args:
        csv_filename: Path of CSV file to prune.
    """
    with open(csv_filename, 'r') as csv_file_obj_r:
        reader = csv.reader(csv_file_obj_r)
        headers = next(reader)
        rows = list(reader)

    # Generated CSV files can have less headers than the rows do entries, so this
    # little blurb pads out the headers on the back to realign things.
    # Otherwise columns are completely missed during processing.
    max_len_row = len(max((row for row in rows), key=len))
    header_len = len(headers)
    for _ in range(max_len_row - header_len):
        headers.append('')

    empty_column_indices = []
    for column_index, header in enumerate(headers):
        for row in rows:
            if row[column_index] != '':
                break
        else:
            if header == '':
                empty_column_indices.append(column_index)

    new_rows = []
    for row in chain([headers], rows):
        new_row = row[:]
        for empty_index in reversed(empty_column_indices):
            new_row.pop(empty_index)
        new_rows.append(new_row)

    with open(csv_filename, 'w') as csv_file_obj_w:
        writer = csv.writer(csv_file_obj_w)
        writer.writerows(new_rows)


def prune_empty_csv_columns(csv_filename: Path, destructive: bool) -> None:
    """Prune empty columns that have a header from a CSV file.
    The user is prompted for every empty column found.

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
            if not destructive:
                print(f'Column "{header}" was found to be empty.')
                choice = obtain_choice_from_user(
                    ['discard', 'keep'], 'Do you wish to {} this column? '
                )
                print()
            else:
                choice = 'discard'

            if choice == 'discard':
                empty_column_indices.append(column_index)

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
    xlsx_filename: Path,
) -> None:
    """Append metadata columns to a CSV from a filename.

    This appends columns with info on the project, page number, and line number to the CSV from a filename.
    If any of these columns already exist, the user will be prompted to keep or overwrite the column based on it's contents.
    Operates in-place on a file.

    Args:
        csv_filename: Path of CSV file to append metadata to.
        filename_regex: Regex to match filenames by.
                        This must be a regex that provides exactly three named groups for the 'project', 'page', and 'line', named as such.
                        This regex will be matched against the original Excel file.
        xlsx_filename: Path of original Excel file CSV was generated from.
                       If given, the filename regex will be applied to this filename, and not the CSV filename.

    Raises:
        KeyError: Provided regex had a disallowed group name.
        ValueError: No fieldnames were found in the provided CSV file.
    """

    match = filename_regex.match(str(xlsx_filename.name))
    if match is None:
        print(
            f'Cannot parse filename "{xlsx_filename}" with given regex /{filename_regex}/. Exiting early without appending metadata'
        )
        return

    metadata = {k.lower(): v for k, v in match.groupdict().items()}

    for key in metadata:
        if key not in ['project', 'page', 'line']:
            raise KeyError(
                f'Filename regex returned a group "{key}" that is not one of "project", "page", or "line"'
            )

    with open(csv_filename, 'r') as csv_file_obj_r:
        reader = csv.DictReader(csv_file_obj_r)
        fieldnames = reader.fieldnames
        if fieldnames is None:
            raise ValueError(f'No fieldnames found in {csv_filename}')
        rows = list(reader)

    new_rows = [row | metadata | {'filename': xlsx_filename.name} for row in rows]
    fieldnames = [*fieldnames, *metadata.keys(), 'filename']

    with open(csv_filename, 'w') as csv_file_obj_w:
        writer = csv.DictWriter(csv_file_obj_w, fieldnames)
        writer.writeheader()
        writer.writerows(new_rows)


def fill_in_missing_entries(csv_filename: Path, enforce_headers: bool) -> None:
    with open(csv_filename, 'r') as csv_file_obj_r:
        reader = csv.DictReader(csv_file_obj_r)
        fieldnames = reader.fieldnames
        if fieldnames is None:
            raise ValueError(f'No fieldnames found in {csv_filename}')
        rows = list(reader)

    if enforce_headers:
        fieldnames = ALLOWED_COLUMN_NAMES

    new_rows = []
    for row in rows:
        new_row = {k: v for k, v in row.items()}
        print(new_row)
        if enforce_headers:
            allowed_headers = ALLOWED_COLUMN_NAMES[:]
            for header in new_row:
                allowed_headers.remove(header)
            for header in allowed_headers:
                new_row[header] = ''

        try:
            if new_row['response'] == '':
                new_row['response'] = 'NR'
        except KeyError:
            pass

        try:
            if new_row['phonetic_transcription'] == '':
                new_row['phonetic_transcription'] = 'see field pages'
        except KeyError:
            pass

        new_rows.append(new_row)

    with open(csv_filename, 'w') as csv_file_obj_w:
        writer = csv.DictWriter(csv_file_obj_w, fieldnames)
        writer.writeheader()
        writer.writerows(new_rows)


def convert_excel_file_to_csvs(
    xlsx_filename: Path,
    sanitize_headers: bool,
    destructive_sanitization: bool,
    discard_empty: bool,
    assume_headers: bool,
    enforce_headers: bool,
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
        destructive_sanitization: Option to, when sanitizing, discard all problems silently.
        assume_headers: Option to assume Excel file first row are column headers.
        prune_empty_columns: Option to prune empty columns.
        append_metadata: Option to append file metadata to the CSV based on its filename. Defaults to False.
        filename_regex: Regex to match filenames by.
                        This must be a regex that provides only three named groups for the project, page and line, named as such.
                        This regex will be matched against CSV files, not Excel files. Remember to account for this if providing alternate regex.
        output_dir: Output directory to place converted files into.
                    The script will create the output directory if it does not exist.
    """

    wb = xl.load_workbook(xlsx_filename, read_only=True)
    for worksheet in wb.worksheets:
        csv_name = output_dir / (
            xlsx_filename.stem.replace(' ', '_') + f'_{worksheet.title}.csv'
        )
        with open(csv_name, 'w') as csv_file_obj:
            writer = csv.writer(csv_file_obj)
            for row in worksheet.values:
                writer.writerow(row)


        strip_csv_whitespace(csv_name)
        prune_empty_csv_rows(csv_name)
        prune_padding_csv_columns(csv_name)
        if prune_empty_columns:
            prune_empty_csv_columns(csv_name, destructive_sanitization)
        if sanitize_headers:
            sanitize_csv_column_names(
                csv_name,
                destructive_sanitization,
                discard_empty,
                assume_headers,
                enforce_headers,
            )
        if append_metadata:
            append_metadata_from_filename(csv_name, filename_regex, xlsx_filename)
        fill_in_missing_entries(csv_name, enforce_headers)


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
    # TODO: Maybe change this to only process CSV in directory made by script this session?
    csv_filenames = list(input_dir.glob('*.csv'))
    merged_filename = csv_filenames[0].name.split('_')[0] + '_merged.csv'

    with ExitStack() as stack:
        csv_files = {
            csv_file: stack.enter_context(open(csv_file, 'r'))
            for csv_file in csv_filenames
            if 'merged' not in csv_file.name
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
    parser = argparse.ArgumentParser(
        description='process and convert LAP Excel files to CSV format. Original Excel files are untouched.'
    )

    group = parser.add_mutually_exclusive_group()
    group.add_argument(
        '-b',
        '--batch',
        action='store_const',
        dest='mode',
        const='batch',
        default='batch',
        help='process whole directory of files (default)',
    )
    group.add_argument(
        '-s',
        '--single-file',
        action='store_const',
        dest='mode',
        const='single',
        help='process single file only',
    )

    parser.add_argument(
        '-c',
        '--no-sanitize-headers',
        action='store_false',
        help='do not sanitize CSV column headers',
    )
    parser.add_argument(
        '-d',
        '--destructive-sanitization',
        action='store_true',
        help='silently discard columns with unsanitary headers without any user prompting. Can cause data to not be converted.',
    )
    parser.add_argument(
        '-j',
        '--discard-empty-columns',
        action='store_true',
        help='silently discard empty columns without any user prompting. Can cause data to not be converted.',
    )
    parser.add_argument(
        '-n',
        '--assume-headers',
        action='store_true',
        help='assume all Excel files have headers rows. Can cause data to not be converted.',
    )
    parser.add_argument(
        '-e',
        '--no-enforce-headers',
        action='store_false',
        help=f'do not enforce at that all column headers are from {'{'}{', '.join(ALLOWED_COLUMN_NAMES)}{'}'}.',
    )
    parser.add_argument(
        '-p',
        '--no-prune-empty-columns',
        action='store_false',
        help='do not prune empty columns from CSV',
    )
    parser.add_argument(
        '-a',
        '--no-append-metadata',
        action='store_false',
        help='do not append filename metadata to CSV',
    )
    parser.add_argument(
        '-f',
        '--filename-regex',
        type=lambda s: re.compile(s),
        default=re.compile(
            r'(?P<project>[a-zA-Z]*)'
            r'_page(?P<page>[a-zA-Z0-9]*)'
            r'line(?P<line>[a-zA-Z0-9]*)'
            r'\.xlsx'
        ),
        help=(
            'expected regex Excel filenames fit under. Must exactly provide named groups for "project", "line", and "page" '
            '(default: (?P<project>[a-zA-Z]*)_page(?P<page>[a-zA-Z0-9]*)line(?P<line>[a-zA-Z0-9]*)\\.xlsx )'
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
        '--output-directory',
        type=Path,
        default=Path('./output/'),
        help='output directory for processed Excel files (default: ./output/ )',
    )
    parser.add_argument(
        'input_path',
        type=Path,
        help='path to Excel (xlsx) file(s) to process',
    )

    return parser


def process_batch(cmd_args: argparse.Namespace) -> None:
    """Process a directory of Excel files into CSV files.

    Args:
        cmd_args: Command line arguments provided by user.

    Raises:
        ValueError: Input path is not a directory.
    """
    if not cmd_args.input_path.is_dir():
        raise ValueError(
            f'Batch mode input path is not a directory: {cmd_args.input_path}'
        )

    for file in cmd_args.input_path.glob('*.xlsx'):
        print(f'Processing: {file}')
        convert_excel_file_to_csvs(
            file,
            sanitize_headers=cmd_args.no_sanitize_headers,
            destructive_sanitization=cmd_args.destructive_sanitization,
            discard_empty=cmd_args.discard_empty_columns,
            assume_headers=cmd_args.assume_headers,
            enforce_headers=cmd_args.no_enforce_headers,
            prune_empty_columns=cmd_args.no_prune_empty_columns,
            append_metadata=cmd_args.no_append_metadata,
            filename_regex=cmd_args.filename_regex,
            output_dir=cmd_args.output_directory,
        )
        if not cmd_args.destructive_sanitization:
            print()

    if cmd_args.merge and not cmd_args.no_sanitize_headers:
        print(
            'Cannot merge without sanitized headers. Finishing without merging any files.'
        )
    elif cmd_args.merge:
        merge_all_csv_in_dir(cmd_args.output_directory, cmd_args.output_directory)


def process_single(cmd_args: argparse.Namespace) -> None:
    """Process a single Excel files into a CSV file.

    Args:
        cmd_args: Command line arguments provided by user.

    Raises:
        ValueError: Input path is not a file.
        ValueError: Input path is not an Excel (.xlsx) file.
    """
    if not cmd_args.input_path.is_file():
        raise ValueError(
            f'Single file mode input path is not a file: {cmd_args.input_path}'
        )
    if cmd_args.input_path.suffix != '.xlsx':
        raise ValueError(
            f'Input path is not an Excel file (.xlsx): {cmd_args.input_path}'
        )

    print(f'Processing: {cmd_args.input_path}')
    convert_excel_file_to_csvs(
        cmd_args.input_path,
        sanitize_headers=cmd_args.no_sanitize_headers,
        destructive_sanitization=cmd_args.destructive_sanitization,
        discard_empty=cmd_args.discard_empty_columns,
        assume_headers=cmd_args.assume_headers,
        enforce_headers=cmd_args.no_enforce_headers,
        prune_empty_columns=cmd_args.no_prune_empty_columns,
        append_metadata=cmd_args.no_append_metadata,
        filename_regex=cmd_args.filename_regex,
        output_dir=cmd_args.output_directory,
    )

    if cmd_args.merge and cmd_args.no_sanitize_headers:
        print(
            'Cannot merge without sanitized headers. Finishing without merging any files.'
        )
    elif cmd_args.merge:
        merge_all_csv_in_dir(cmd_args.output_directory, cmd_args.output_directory)


def main() -> None:
    """Process Excel files into CSV files.

    Raises:
        FileNotFoundError: Input directory does not exist.
        OSError: Error occurred making output directory (if it did not exist).
    """
    parser = build_arg_parser()
    args = parser.parse_args()

    if not args.input_path.exists():
        raise FileNotFoundError('Input path does not exist')

    if not args.output_directory.exists():
        try:
            args.output_directory.mkdir(exist_ok=True, parents=True)
        except OSError as e:
            raise OSError('Issue occurred making output directory') from e

    if args.mode == 'batch':
        process_batch(args)

    elif args.mode == 'single':
        process_single(args)


if __name__ == '__main__':
    main()

from os import getcwd
from os import walk
from pandas import read_excel
from pandas import concat
from pandas import DataFrame
from pandas import ExcelWriter
from numpy import ceil
from numpy import floor

OPTIONS = {
    0: 'None',
    1: 'Differentiation',
    2: 'Supplementary',
    3: 'Substantial',
    4: 'Extensive'
}

BULLET = '\u2022'


def round_df(df):
    """ df is a one-dimensional dataframe"""
    for row, entry in enumerate(df):
        if entry >= ceil(entry) - 0.5:
            # Round up
            df[row] = ceil(entry)
        else:
            # Round down
            df[row] = floor(entry)


def clean_name(name):
    """ Takes a name and returns it in format: 'First Last'"""
    # Remove trailing whitespace
    name = name.strip()
    # Separate first and last names
    names = name.split()
    # Insurance against more than two words
    first = capitalise(names[0])
    last = capitalise(names[-1])
    return f'{first} {last}'


def capitalise(name):
    """ Capitalises name"""
    FIRST_INITIAL = 0
    formatted_name = ''
    for i, char in enumerate(name):
        if char.isalpha():
            # Is an alphanumeric value
            if i == FIRST_INITIAL:
                # First letter in name
                formatted_name += char.upper()
            else:
                formatted_name += char.lower()
    return formatted_name


def frequency_map(name):
    """ Gets a frequency map of the name"""
    freq_map = {}
    for c in name:
        if c.isalpha():
            # Is an alphanumeric value
            # Convert to uppercase
            upper_c = c.upper()
            if upper_c not in freq_map:
                # Add character to map
                freq_map[upper_c] = 1
            else:
                # Increment count for character
                freq_map[upper_c] += 1
    return freq_map


def name_similarity(name_one, name_two):
    """ Returns the similarity of the two names """
    # Get frequency map for both names
    name_one_map = frequency_map(name_one)
    name_two_map = frequency_map(name_two)
    num_differences = 0
    for letter in name_one_map:
        difference = abs(
            name_one_map[letter] - name_two_map.get(letter, 0)
        )
        num_differences += difference
    # Find the largest name
    min_name_size = min(len(name_one), len(name_two))
    return 100 * (1 - num_differences / min_name_size)


def similarity_check(names):
    """ Iterates through a list of names and
    checks for any inconsistencies"""
    names = list(names)
    ignore = []
    replaced = []
    # Ask user for desired similarity
    while True:
        try:
            min_similarity = int(
                input('Please enter similarity threshold (%): ')
            )
        except ValueError:
            # Failed to convert to integer
            print('Error: please enter a whole number')
            continue
        if min_similarity <= 0 or min_similarity >= 100:
            # Not a valid percentage
            print('Error: please enter a number between 0 and 100')
            continue
        else:
            # valid percentage
            break
    # Check for words which satisfy desired similarity
    for index_one, name_one in enumerate(names):
        for index_two, name_two in enumerate(names[index_one:]):
            if name_one != name_two:
                # Two names are not equal
                if (name_one, name_two) in ignore \
                        or (name_two, name_one) in ignore:
                    # Ignore names
                    continue
                elif name_one in replaced or name_two in replaced:
                    # Replaced names
                    continue
                similarity = name_similarity(name_one, name_two)
                if similarity >= min_similarity:
                    # Two names are similar
                    print(
                        f'\nGreater than {min_similarity}% similarity found '
                        f"between names '{name_one}' and '{name_two}'.\n\n"
                        f'What name would you like to replace:\n\n'
                        f"Option 1. '{name_one}' with '{name_two}'\n"
                        f"Option 2. '{name_two}' with '{name_one}'\n"
                        f'Option 3. no change\n'
                    )
                    option = ''
                    choices = ['1', '2', '3']
                    while option not in choices:
                        option = input('Enter option: ')
                        if option not in choices:
                            # Not a valid option
                            print(
                                'Error: please choose an option between '
                                '1, 2 and 3!'
                            )
                    min_index = min(index_one, index_two)
                    if option == '1' or option == '2':
                        for index, name in enumerate(names[min_index:], min_index):
                            if name == name_one and option == '1':
                                # Change name to name_two
                                names[index] = name_two
                            elif name == name_two and option == '2':
                                # Change name to name_one
                                names[index] = name_one
                        if option == '1':
                            # Replaced name_one
                            replaced.append(name_one)
                        else:
                            # Replaced name_two
                            replaced.append(name_two)
                    else:
                        # Add to ignore list
                        ignore.append((name_one, name_two))
    return names


def import_sheets(app_path):
    """ Find all .xlsx files in a given folder. Returns a dataframe containing
    the raw data along with an additional column for the Sheet name."""
    all_sheets = []
    while not all_sheets:
        folder_name = input('Enter folder name: ')
        # Code from https://stackoverflow.com/questions/3207219/
        file_names = next(
            walk(f'{app_path}\\{folder_name}'),
            (None, None, [])
        )[2]
        if not file_names:
            # No files in location
            print(f'Either {folder_name} is empty, or does not exist!')
            continue
        for file in file_names:
            if file.endswith('.xlsx'):
                # Code from https://stackoverflow.com/questions/16888888/
                sheet = read_excel(f'{folder_name}\\{file}')
                # Add column for sheet name
                sheet['Sheet'] = file
                all_sheets.append(sheet)
        if not all_sheets:
            # No Excel sheets in location
            print(f'No Excel (.xlsx) sheets in folder {folder_name}!')
    return folder_name, concat(all_sheets, ignore_index=True)


def import_column_names():
    """ Creates a list of column names from which data should be extracted,
    using the contents of a given .txt file.
    The first line of the text file should contain the name of column against
    which data will be sorted and grouped."""
    column_names = []
    while not column_names:
        file_name = input('Enter name of column text file: ')
        # Try opening the text file
        try:
            file = open(f'{file_name}.txt', 'r')
        except IOError:
            print(f'{file_name}.txt does not exist in current directory!')
            continue
        # Succeeded in opening the text file
        column_names = file.readlines()
        # Check to see if text file is empty
        if not column_names:
            print(f'{file_name}.txt is empty!')
    # Strip any white-space or newline characters
    for index, name in enumerate(column_names):
        column_names[index] = name.strip(' \n')
    return file_name, column_names


def check_names(total_df, first_column):
    """ Asks the user whether they want to check provided names for
    misspellings."""
    check_names = ''
    while check_names not in ['y', 'n']:
        check_names = input(
            'Would you like to check for student name errors [y/n]: '
        )
        if check_names.isalpha():
            check_names = check_names.lower()
        if check_names == 'y':
            # Look for errors
            correct_names = similarity_check(total_df[first_column])
            # Replace first column with corrected names
            total_df[first_column] = DataFrame(
                correct_names,
                columns=[first_column]
            )


def proceed_prompt():
    """ Asks the user if they want to proceed or not."""
    keep_going = -1
    while keep_going not in ['y', 'n']:
        keep_going = input('Would you like to process more data [y/n]: ')
        if keep_going.isalpha():
            keep_going = keep_going.lower()
    if keep_going == 'n':
        # Stop
        exit()


def value_check(df, sheet_df):
    """ Checks for invalid entries in numerical column."""
    violations = []
    for row, entry in enumerate(df):
        sheet = sheet_df[row]
        if not isinstance(entry, int) and not isinstance(entry, float):
            # Not a numerical type
            violations.append(
                f"{BULLET} Entry '{entry}' in column '{df.name}' in sheet"
                f" '{sheet}' is not a numeric type!"
            )
        elif entry not in OPTIONS:
            # Not a valid entry
            violations.append(
                f"{BULLET} Entry '{entry}' in column '{df.name}' in sheet"
                f" '{sheet}' is not a valid entry!"
            )
    return violations


def main():
    """Handles user interaction"""
    app_path = getcwd()
    while True:
        # Import data from sheets in given folder and collate in dataframe
        folder_name, total_df = import_sheets(app_path)
        # Import list of column headers
        file_name, column_names = import_column_names()
        # Drop any columns in dataframe which are not in column_names
        # (excluding the Sheets column)
        for col in total_df:
            if col not in column_names and col != 'Sheet':
                total_df.drop(col, inplace=True, axis=1)
        # Check for empty dataframe
        if total_df.empty:
            print(
                f'Error: sheets in folder {folder_name} do not contain any of'
                f' the columns specified in {file_name}.txt!'
            )
            continue
        # Use first column to sort (assumed to be names)
        first_column = column_names[0]
        # Remove NaNs in names
        total_df = total_df[total_df[first_column].notnull()]
        # Clean up names (if necessary)
        total_df[first_column] = total_df[first_column].apply(clean_name)
        # Keep track of entry violations
        total_violations = []
        for col in total_df:
            if col != 'Other' and col != 'Sheet' and col != first_column:
                # Replace attribute NaNs with 0 (None)
                total_df[col] = total_df[col].fillna(value=0)
                # Check for column violations
                total_violations.extend(
                    value_check(total_df[col], total_df['Sheet'])
                )
        num_violations = len(total_violations)
        if num_violations > 0:
            # Some violations occurred
            print(
                f'Error: {num_violations} sheet entry violation(s) was/were'
                ' found, as follows:'
            )
            # Print violations
            for violation in total_violations:
                print(violation)
            continue
        # Reset index
        total_df = total_df.reset_index()
        # Drop index columns
        total_df.drop('index', inplace=True, axis=1)
        # Check for misspelled names
        check_names(total_df, first_column)
        # Sort by names
        total_df = total_df.sort_values(by=first_column)
        # Group entries by names
        grouped_df = total_df.groupby(first_column).mean()
        # Add columns for means
        mean_df = grouped_df.mean(axis=1)
        # Round grouped_df and mean_df
        for col in grouped_df:
            if col != first_column:
                round_df(grouped_df[col])
        # Round mean_df
        round_df(mean_df)
        # Try to round grouped_df and mean_df to integer
        try:
            grouped_df = grouped_df.astype('int64')
            mean_df = mean_df.astype('int64')
        except ValueError:
            # Found some NaN values
            print(
                f'Error: failed to find any columns in {file_name}.txt'
                ' which contain numeric data!'
            )
            continue
        assessment = []
        for index, entry in enumerate(mean_df, 1):
            # Convert from integer to option
            assessment.append(OPTIONS[entry])
        assessment_df = DataFrame({'Mean': assessment}, index=grouped_df.index)
        combined_df = concat([grouped_df, assessment_df], axis=1)
        # Code from https://lifewithdata.com/2022/02/17/pandas-to_excel-write
        # -a-dataframe-to-an-excel-file/
        with ExcelWriter(f'{app_path}\\Processed Data - {folder_name}.xlsx') \
                as writer:
            total_df.to_excel(writer, sheet_name='combined', index=False)
            combined_df.to_excel(writer, sheet_name='average', index=True)
        # Determine if user wants to process more data
        proceed_prompt()


if __name__ == "__main__":
    main()

print("Loading program...")

import json
import logging
import os
import sys
from datetime import datetime as dt
from inspect import cleandoc
from os.path import dirname, exists, join, split
from textwrap import dedent
from time import perf_counter

import pandas as pd
from dateutil.relativedelta import relativedelta

from dbmerger.data import Data
from dbmerger.export import Export
from dbmerger.settings import Settings

now = dt.strftime(dt.now(), "%Y-%m-%d_%H-%M-%S")
log_path = join(dirname(__file__), 'log')
log_file = join(log_path, f'log_{now}.txt')

if not exists(log_path):  # if log folder doesn't exist
    os.makedirs(log_path)  # create log folder
elif len(os.listdir(log_path)) > 30:  # keep files no older than 2 months
    for file in sorted(os.listdir(log_path)):
        file_dt = dt.strptime(file, "log_%Y-%m-%d_%H-%M-%S.txt")
        if file_dt < dt.now() - relativedelta(months=2):
            os.remove(join(log_path, file))

logging.basicConfig(filename=log_file,
                    filemode='w',
                    format='%(asctime)s: [%(lineno)d: %(module)s.%(funcName)s] - %(message)s ',
                    datefmt='%y-%b-%d %H:%M:%S')
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
stdout_handler = logging.StreamHandler(stream=sys.stdout)
stdout_handler.setLevel(logging.INFO)
logger.addHandler(stdout_handler)


def handle_exception(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return

    logger.critical("CRITICAL ERROR: Uncaught Exception", exc_info=(exc_type, exc_value, exc_traceback))


sys.excepthook = handle_exception
logger.debug("Program loaded")
start_time = perf_counter()

print("""
Database Merge & Compare

Version: 1.0
Date: 18/06/21
Author: George M. Marino""")

print("""
Matches customer databases to the primary database and outputs excel files containing 
matches based on pre-defined match conditions.
Able to take many different customer excel sheets, outputting matches on a 
case-by-case basis simultaneously.
Input/output folder and other settings configurable in the settings.txt file 
found in the program directory.
""")
print('-' * 88, '\n')


class Match(Data, Export):
    """Import and merge databases"""

    def __init__(self, settings):
        Data.__init__(self, settings)
        Export.__init__(self, settings)

        # options
        self.options = settings['options']

        # columns names
        self.prim_cols = settings['prim_cols']
        self.prim_id = settings['id']
        self.prim_areas = settings['areas']

        self.cust_cols = settings['cust_cols']
        self.cust_areas = settings['cust_areas']

        # directories
        self.check_path = settings['check']

        # system specific open command
        self.open = settings['open']

        # create directories as needed
        if not os.path.exists(self.check_path):
            logging.debug(f'Creating check path at {self.check_path}')
            os.makedirs(self.check_path)

        logger.debug(f'========== Match object instantiated ==========\n{json.dumps(self.__dict__, indent=2)}')

    def get_matches(self, primary, cust, prim_main, cust_main):
        """Match by name, iterate through each match condition, and generate separated dataframes"""
        logger.info('========== Matching ==========')
        print()

        try:  # generate matches by name
            logger.info('Matching by name... ')
            time = perf_counter()
            cust_name = [col for col in self.cust_cols['n'] if col in cust.columns]

            if len(cust_name) == 0:
                return None, None, None

            matches = pd.merge(primary, cust, left_on=self.prim_cols['n'], right_on=cust_name,
                               how='inner').drop(cust_name, axis=1)
            logger.info(f'Done: {round(perf_counter() - time, 5)}s')
            print()
        except KeyError:
            return None, None, None

        # empty sets for storing indices
        same_idx = set()
        diff_idx = set()
        none_idx = set()

        # filter and drop column types not being used for matching
        logger.debug('Filtering column dictionaries')
        prim_cols = {k: list(set(v).intersection(set(primary.columns)))
                     for k, v in self.prim_cols.copy().items() if k in self.options}
        cust_cols = {k: list(set(v).intersection(set(cust.columns)))
                     for k, v in self.cust_cols.copy().items() if k in self.options}

        for (i, prim_col), (_, cust_col) in zip(prim_cols.items(), cust_cols.items()):
            if i == 'n':  # skip over checking names
                continue
            text = {'p': 'Phone number', 'e': 'Email', 'f': 'Fax number'}

            if len(cust_col) == 0:  # if columns not in customer df, skip
                logger.warning(f'{text[i]} columns not found from customer database, skipping...')
                continue

            logger.debug(f'{text[i]} columns found: {cust_col}')

            # update indices sets

            logger.info(f'Matching by {text[i].lower()}... (found {len(cust_col)} customer column/s)')
            time_ = perf_counter()
            same_idx, diff_idx, none_idx = self.get_match_idx(matches, prim_col, cust_col,
                                                              same_idx, diff_idx, none_idx)
            logger.info(f'Done matching all {text[i].lower()}s: {round(perf_counter() - time_, 5)}s\n')

        # return filtered matches dataframes
        same = matches[matches[self.prim_id].isin(same_idx)].drop_duplicates(subset=self.prim_id)
        diff = matches[matches[self.prim_id].isin(diff_idx)].drop_duplicates(subset=self.prim_id)
        none = matches[matches[self.prim_id].isin(none_idx)].drop_duplicates(subset=self.prim_id)

        # filter out entries where customer data has already been matched
        diff = diff[~diff['cust_index'].isin(same['cust_index'])]
        none = none[~none['cust_index'].isin(same['cust_index'])]

        # match uncertain matches by areas
        same, diff, none = self.match_areas(same, diff, none, matches, prim_main, cust_main)

        logger.debug(f'Matches - same: {len(same)}, diff: {len(diff)}, none: {len(none)}')
        logger.info('========== Matching Complete ==========\n')
        if all(char not in self.options for char in ['x', 's']) or 'v' in self.options:
            print('-' * 88, '\n')

        return same, diff, none

    def get_match_idx(self, matches, prim_col, cust_col, same_idx, diff_idx, none_idx):
        """Iterate through given customer columns checking for matches in given primary database columns
        
        :return: sets for perfect matches, uncertain matches, and missing info available"""
        for col in cust_col:  # iterate through customer columns
            logger.info(f'Searching {col} column... ')

            # drop already found perfect matches
            logger.debug('Dropping matches already found')
            match_drop = matches[~matches[self.prim_id].isin(same_idx)]
            logger.debug(f'New match_df shape: {match_drop.shape}')

            # perfect matches
            same = match_drop[match_drop[prim_col].isin(match_drop[col]).any(1)]

            # uncertain matches
            diff = match_drop[(~match_drop[prim_col].isin(match_drop[col]).any(1)) &
                              (~match_drop[[col] + prim_col].isna().all(1))]

            # missing data
            none = match_drop[match_drop[[col] + prim_col].isna().all(1)]

            # extract primary dataframe indices
            same_idx = same_idx | set(same[self.prim_id])
            diff_idx = (diff_idx | set(diff[self.prim_id])) - same_idx
            none_idx = (none_idx | set(none[self.prim_id])) - same_idx - diff_idx

            logger.debug(f'Current matches - same: {len(same)}, diff: {len(diff)}, none: {len(none)}')

        return same_idx, diff_idx, none_idx

    def match_areas(self, same, diff, none, matches, prim_main, cust_main):
        """match based on medical areas"""
        start_lengths = [len(same), len(diff), len(none)]

        # extract available columns
        prim_cols = set(self.prim_areas.keys()).intersection(prim_main.columns)
        cust_cols = set(self.cust_areas).intersection(cust_main.columns)
        cust_df = cust_main[cust_cols].reset_index().rename(columns={'index': 'cust_index'}).fillna('')

        if len(prim_cols) == 0:
            logger.warning(f'No medical area columns found in primary database or areas.txt empty, skipping areas check...\n')
            return same, diff, none

        if len(cust_cols) == 0:
            logger.warning(f'No medical area columns found in customer database or areas.txt empty, skipping areas check...\n')
            return same, diff, none

        logger.info(f'Checking medical areas for imperfect matches... (found {len(cust_cols)} customer column/s)')
        logger.debug(f'Filtered medical area columns: {cust_cols}')
        time = perf_counter()

        logger.debug(cleandoc(f"""Start lengths -
                     same: {start_lengths[0]},
                     diff: {start_lengths[1]},
                     none: {start_lengths[2]}""").replace('\s+', ' ').replace('\n', ''))

        alt_areas = {k: v for k, v in self.prim_areas.items() if len(v) > 0}

        # convert encoded columns to one column of lists
        logger.debug('Converting encoded columns to series of lists')
        prim_areas = prim_main[[self.prim_id, *prim_cols]].set_index(self.prim_id)
        prim_areas = prim_areas.apply(lambda row: [i * col for i, col in zip(row, prim_cols) if i != 0], axis=1)
        prim_areas = prim_areas.rename('Primary DF Areas').to_frame().reset_index()

        def search_areas(row):
            """return True or False if match found in customer column"""
            match = False
            for cust in cust_cols:  # iterate through available cust_cols
                for area in row['Primary DF Areas']:
                    match = area in row[cust]
                    if len(alt_areas.get(area, [])) > 0:
                        match = any([match, any([alt in row[cust] for alt in alt_areas[area]])])

                    if match:  # early break if found
                        return match
            return match

        # merge and reorder columns
        logger.debug('Merging with diff / none')
        diff = diff.merge(prim_areas, on=self.prim_id).merge(cust_df, on='cust_index')
        none = none.merge(prim_areas, on=self.prim_id).merge(cust_df, on='cust_index')

        # extract boolean matching for indexing all match cases
        logger.debug('Matching areas')
        diff_bool = diff.apply(lambda x: search_areas(x), axis=1)
        none_bool = none.apply(lambda x: search_areas(x), axis=1)

        logging.debug('Building match dataframes for each case')
        # extract IDs for same condition from both given and new matches
        same_idx = {*same[self.prim_id], *diff[diff_bool][self.prim_id], *none[none_bool][self.prim_id]}
        # index matches for same condition
        same = matches[matches[self.prim_id].isin(same_idx)].drop_duplicates(subset=self.prim_id)
        # add extra matching data for areas and index remaining data for diff / none
        same = same.merge(prim_areas, on=self.prim_id).merge(cust_df, on='cust_index')
        diff = diff[~diff_bool]
        none = none[~none_bool]

        # reformat lists to single strings
        same['Primary DF Areas'] = same['Primary DF Areas'].apply(lambda x: ', '.join(x))
        diff['Primary DF Areas'] = diff['Primary DF Areas'].apply(lambda x: ', '.join(x))
        none['Primary DF Areas'] = none['Primary DF Areas'].apply(lambda x: ', '.join(x))

        logger.debug(cleandoc(f"""Differences - 
                     same: +{len(same) - start_lengths[0]}, 
                     diff: {len(diff) - start_lengths[1]}, 
                     none: {len(none) - start_lengths[2]}""").replace('\s+', ' ').replace('\n', ''))
        logger.info(f'Done: {round(perf_counter() - time, 5)}s\n')

        return same, diff, none

    def conflicts(self, filename, same, diff, none):
        """Show conflicts to user and allow them to view more information"""

        # concat dataframes
        df = pd.concat([diff, none])
        if len(df) == 0:  # skip if none
            logger.debug(f'No conflicts, skipping conflicts function\n')
            return df

        if 'q' not in self.options:
            logger.info('========== Resolve Conflicts ==========')
            print()
        else:
            logger.debug('========== Resolve Conflicts ==========')

        view = False

        if 's' not in self.options:
            # prompt to view more information
            logger.info(f"Found {len(df)} conflicts")
            print(dedent("""
            How would you like to resolve these conflicts?
            - x: export conflicts to check folder and delete rows you would like to remove
            - v: view more information on conflicts in program and manually type in IDs to keep
            """))
            view = 'v' in input("Enter 'x' or 'v': ").lower().strip()
            logger.debug(f'User input for view in console: {view}')

            print('\n' + '-' * 88 + '\n')

        # resolve conflicts
        if 'q' in self.options:
            if 'o' in self.options:
                logger.info('Auto-mode enabled: only keeping perfect matches')
                keep = []
            else:
                logger.info('Auto-mode enabled: automatically keeping all conflicts')
                keep = set(df[self.prim_id])
        elif view or ('v' in self.options and 's' in self.options):
            conflicts = sorted(list(set(df[self.prim_id])))
            self.view_conflicts(df, conflicts)
            keep = self.resolve_conflicts_by_input(df, conflicts)
            print()
        else:  # resolve conflicts by spreadsheet
            if 'x' not in self.options:
                print()
                self.options += ['x']
                self.check(filename, same, diff, none)
            keep = self.resolve_conflicts_by_check(filename)
            print()

        logger.info(f"======== Keeping {len(keep)} conflicting entries ========\n")

        # filter dataframe
        return df[df[self.prim_id].isin(keep)]

    def view_conflicts(self, df, conflicts):
        """view information on conflicts in program window"""
        logger.debug('Resolving conflicts in console')

        while True:
            logger.debug('Prompting user for entries to view in console')
            # show IDs with conflicts and prompt user for IDs to view
            print(f"IDs with conflicts: ")
            [print(conflicts[i:i + 15]) for i in range(0, len(conflicts), 15)]

            print("\nEnter IDs you would like to check separated by a space")
            print("---OR---")
            print("Enter 'n' to stop viewing information\n""")

            inp = input("(e.g. '2 8 24 849' or 'all') : ").lower().strip()
            logger.debug(f'Input: {inp}')

            if inp == 'n':  # break loop if user inputs 'n'
                print('\n', '-' * 28, '\n', sep='')
                break
            elif inp == 'all':  # display all conflicts
                inp = list(df[self.prim_id])

            try:  # recast as integers
                idx = [int(i) for i in inp.split(' ')]
                logger.debug(f'Output: {idx}')
            except ValueError:  # error if non-numbers entered
                print()
                logger.error('!!! ERROR: only enter numbers separated by spaces !!!\n')
                continue
            except AttributeError:  # return all indices
                logger.debug('Attribute Error: returning all indices\n')
                idx = inp

            try:
                logger.debug(f'Indices: {idx}')
                logger.debug('Printing entries in console')
                for i in idx:  # print series for each index
                    print('\n', '-' * 28, '\n', sep='')
                    print(df.set_index(self.prim_id).loc[i])
            except KeyError:  # indices given not in conflicts
                print()
                logger.error('!!! ERROR: index not found, enter only numbers from the conflicts list !!!\n')
                continue

            print('\n', '-' * 28, '\n', sep='')

    def resolve_conflicts_by_input(self, df, conflicts):
        """get IDs to keep from user in program window"""
        while True:
            logger.debug(f'Prompting user for entries to keep')

            # show IDs with conflicts and prompt user for IDs to keep
            print(f"IDs with conflicts: ")
            [print(conflicts[i:i + 15]) for i in range(0, len(conflicts), 15)]
            print("\nEnter IDs you would like to keep")
            inp = input("(e.g. '6 12 746' or 'all' or 'none') : ").lower().strip()
            logger.debug(f'Input: {inp}')

            if inp == 'all':  # keep all IDs
                return list(df[self.prim_id])
            elif inp == 'none':  # keep no IDs
                return []
            else:
                try:  # recast as integers
                    keep = [int(i) for i in inp.split(' ')]
                    logger.debug(f'Output: {keep}')
                    if set(keep).issubset(set(df[self.prim_id])):  # if all IDs given are conflict IDs
                        logger.debug('All input entries successfully found')
                        return keep
                    else:  # retry
                        print()
                        logger.error('!!! ERROR: some IDs not found in conflict list !!!\n')
                        print()
                except ValueError:  # error if non-numbers entered
                    print()
                    logger.error('!!! ERROR: only enter numbers separated by spaces !!!\n')
                    print()
                    continue
                except Exception as e:  # other errors
                    logger.error(f'{e}\n')

    def resolve_conflicts_by_check(self, filename):
        """load IDs to keep from check spreadsheet"""
        logger.debug(f'Resolving conflicts by check spreadsheet')
        logger.debug(f'Filename: {filename}_check.xlsx')

        check_name = join(self.check_path, f'{filename}_check.xlsx')  # full filepath

        path = check_name
        folders = []
        while True:  # split filepath
            path, folder = split(path)
            if folder != "":
                folders.append(folder)
            elif path != "":
                folders.append(path)
                folders.reverse()
                break

        short_name = '/'.join(folders[-3:])  # shortened filepath for printing

        # run system specific open command
        logger.debug(f'Opening file for platform: {sys.platform}')
        if sys.platform == "win32":
            os.startfile(check_name)
        else:
            os.system(f'{self.open}{check_name}')

        # wait for user to continue
        print(cleandoc(f"""From your desktop, open {short_name}
        Delete rows you don't want to keep from the 'Uncertain Match' and 'Missing Data' sheets
        
        !!! Remember to save the spreadsheet before continuing !!!"""), '\n')

        input("Hit enter when you are ready to load from check file: ")

        logger.debug('Loading entries from spreadsheet')

        # extract remaining IDs
        check_diff = list(pd.read_excel(check_name, sheet_name='Uncertain Match')[self.prim_id])
        check_none = list(pd.read_excel(check_name, sheet_name='Missing Data')[self.prim_id])
        logger.debug(f'Keeping - diff: {len(check_diff)}, none: {len(check_none)}')
        return check_diff + check_none

    def loop_customer_files(self):
        """Loop through all customer excel sheets in input folder"""
        # convert excel files to csv
        self.convert_excel()

        # get filenames
        prim_filename, cust_filenames = self.check_files()
        if prim_filename is None:  # stop program if no files in input folder
            logger.debug('prim_filename is None')
            return False

        # load and process primary df
        prim_main, prim_copy = self.get_df(prim_filename, kind='primary')

        for filename in cust_filenames:  # loop through each customer df
            logger.info('=' * 88)
            logger.info(f'BEGIN PROCESSING: {filename}')
            logger.info('=' * 88 + '\n')

            try:  # load customer df
                cust_main, cust_copy = self.get_df(filename, kind='cust')
            except KeyError:  # skip if given column names not found in df
                logger.error(f'ERROR: Skipping {filename} - columns not found')
                print('Change column names in customer database to match program settings')
                print('or run the program again with different column name settings\n')
                continue

            # filename for export
            filename = filename.replace('.csv', '')
            logger.debug(f'Running matches for {filename}')

            # get matches
            same, diff, none = self.get_matches(prim_copy, cust_copy, prim_main, cust_main)
            if same is None:
                logger.error(f'ERROR: Skipping {filename} - columns not found')
                print('Change column names in customer database to match program settings')
                print('or run the program again with different column name settings\n')
                continue

            # export 3 matching conditions to excel sheet
            if 'x' in self.options:
                self.check(filename, same, diff, none)

            # resolve conflicts
            keep = self.conflicts(filename, same, diff, none)

            # concat all matches
            final_matches = pd.concat([same, keep])

            if len(final_matches) == 0:
                logger.info(f'No Matches Found: Skipping final export and continuing to next file\n')
                continue

            logger.debug('========== Final Exports ==========')
            # export final matches
            self.final(filename, final_matches, prim_main, cust_main)

            # export mailing list
            if 'l' in self.options:
                self.mail(filename, final_matches, prim_main, cust_main, kind='perfect')

            # add uncertain mailing lists
            if 'u' in self.options:
                check_name = join(self.check_path, f'{filename}_check.xlsx')  # full filepath
                diff = pd.read_excel(check_name, sheet_name='Uncertain Match')
                none = pd.read_excel(check_name, sheet_name='Missing Data')
                conflict_matches = pd.concat([diff, none])
                logger.debug(f'Processing {len(conflict_matches)} conflicting matches for mailing list')
                self.mail(filename, conflict_matches, prim_main, cust_main, kind='uncertain')

            logger.info('========== Final Exports Complete ==========\n')

        return True


all_settings = Settings().get_all()
logger.debug(f'Final settings\n{json.dumps(all_settings, indent=2)}')

# noinspection PyBroadException
try:
    main = Match(all_settings)
    open = main.loop_customer_files()
except Exception:
    logger.exception("CRITICAL ERROR: Caught Exception")
    open = True

if open:
    logger.debug('Opening final folder path')
    if sys.platform == "win32":
        os.startfile(all_settings['final'])
    else:
        os.system(f"{all_settings['open']}{all_settings['final']}")

logger.info(f'Done. Total Runtime: {perf_counter() - start_time}s')
input('Press return to exit.')

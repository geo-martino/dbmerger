import glob
import json
import logging
import os
import re
import shutil
import sys
from itertools import chain
from multiprocessing import cpu_count
from multiprocessing.dummy import Pool
from os.path import basename, dirname, join
from time import perf_counter

import numpy as np
import pandas as pd

import dbmerger.xlsx2csv as xlsx2csv

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
stdout_handler = logging.StreamHandler(stream=sys.stdout)
stdout_handler.setLevel(logging.INFO)
logger.addHandler(stdout_handler)


class Data:

    def __init__(self, settings):

        # options
        self.options = settings['options']

        # columns names
        self.prim_cols = settings['prim_cols']
        self.prim_id = settings['id']

        self.cust_cols = settings['cust_cols']

        # mailing list/phone options
        self.phone_drop = settings['phone_drop']

        # directories
        self.input_path = settings['input']
        self.prim_filename = settings['prim_filename'] + '.csv'

        # create directories as needed        
        if not os.path.exists(self.input_path):  # if input folder doesn't exist
            logging.debug(f'Creating input path at {self.input_path}')
            os.makedirs(self.input_path)  # create input folder

        logger.debug(f'========== Data object instantiated ==========\n{json.dumps(self.__dict__, indent=2)}')

    def check_files(self):
        """get filenames and check for primary database presence"""
        # get input database filenames
        cust_filenames = [basename(file) for file in glob.glob(f'{self.input_path}/*.csv')
                          if not file.startswith('~$')]
        logger.debug(f'Files in input folder:\n{cust_filenames}')

        # no csv files in folder
        if len(cust_filenames) == 0:
            logger.critical("Input folder doesn't contain any csv files!")
            print("Add your files to the following path and restart the program:")
            print(self.input_path, '\n')
            return None, None

        # check for primary database presence
        while self.prim_filename not in cust_filenames:
            logger.warning(f"Primary Database not found! (Filename {self.prim_filename} not found in input folder)")
            logger.debug('Getting user input')
            print()
            logger.info('Input folder contains the following csv files: ')
            [logger.info(file) for file in cust_filenames]  # print list of csv files
            print(f"\nDon't see your file? Move it to {self.input_path} and restart the program!\n")

            # get filename from user
            name = input('Enter filename for Primary database (case-sensitive) : ')
            logger.debug(f'input: {name}')
            name = re.sub(r'[\\/*?:"<>|]', '', name)
            if name == '':  # use default if not given
                name = 'prim_database'

            if not name.endswith('csv'):  # append file type if not given
                self.prim_filename = name + '.csv'
            logger.debug(f'output: {self.prim_filename}')
            print('\n', '-' * 28, '\n', sep='')

        # check again for file conversion if needed
        self.convert_excel()
        # drop primary database from customer database list
        cust_filenames.remove(self.prim_filename)
        logger.debug(f'Primary Filename: {self.prim_filename}')
        logger.debug(f'Customer Files:\n{cust_filenames}\n')
        return self.prim_filename, cust_filenames

    @staticmethod
    def drop_dupes(df, cols):
        """attempts to drop duplicates based on matching info by data type keep rows containing most info"""

        # function to check for dupes for numpy apply method
        def check_dupes(data):
            # boolean indexing for NaNs
            # must compare NaN values to string 'nan' for string arrays
            if not isinstance(data[0], str):
                nans = np.isnan(arr)
            else:
                nans = (arr == np.array([np.nan]).astype(str)[0])

            # boolean indexing for matches
            matches = (arr == data)
            rows = np.where((matches + nans).all(1) * ~nans.all(1))[0]

            # if duplicates and not already logged
            if len(rows) > 1 and all(row not in drop for row in rows):
                # boolean indexing for NaNs
                # must compare NaN values to string 'nan' for string arrays
                def nunique(a):
                    # return n-unique values in one row of array of integers or strings
                    if not isinstance(a[0], str):
                        return len(set(a[~np.isnan(a)]))
                    else:
                        return len(set(a[a != np.array([np.nan]).astype(str)[0]]))

                # row with most info is kept 
                count = np.apply_along_axis(nunique, 1, arr)
                drop.extend([v for k, v in zip(count, rows) if k != max(count)])

                return [False] * len(data)
            return [True] * len(data)

        # build duplicates to drop for each data type
        drop = []
        text = {'p': 'phone numbers', 'e': 'emails', 'f': 'fax numbers'}
        for i, col in cols.items():
            if i not in text.keys() or len(col) == 0:
                continue

            logger.info(f'Dropping duplicates for {text[i]}... ')

            if i == 'e':
                arr = df[col].to_numpy(dtype=str)
                np.apply_along_axis(check_dupes, 1, arr)
            elif i != 'n':
                arr = df[col].to_numpy(dtype=float)
                np.apply_along_axis(check_dupes, 1, arr)

        drop = list(set(drop))  # get unique values
        logger.info(f'Deleted {len(drop)} entries containing duplicate data')
        # [print(drop[i:i + 15]) for i in range(0, len(drop), 15)]  # print row numbers
        return df[~df.index.isin(drop)]  # drop duplicates

    def clean_df(self, df, kind):
        """copy and clean dataframe"""
        if kind == 'primary':  # primary columns
            cols = {k: list(set(v).intersection(set(df.columns))) for k, v in self.prim_cols.copy().items()}
            nme = cols.get('n', [])
            cols['n'] = [self.prim_id] + cols['n']
            all_col = list(chain(*cols.values()))
        else:  # customer columns
            cols = {k: list(set(v).intersection(set(df.columns))) for k, v in self.cust_cols.copy().items()}
            nme = cols.get('n', [])
            all_col = ['cust_index'] + list(chain(*cols.values()))

        num = cols.get('p', [])
        eml = cols.get('e', [])
        fax = cols.get('f', [])

        logger.debug(f'Filtered names: {nme}')
        logger.debug(f'Filtered phones: {num}')
        logger.debug(f'Filtered emails: {eml}')
        logger.debug(f'Filtered fax: {fax}')
        logger.debug(f'Filtered all: {all_col}')

        # drop customers who don't want to be contacted
        if 'i' in self.options and kind == 'primary':
            df = df[df['Interesse MAFO'] != 2]
            logger.debug('Dropping customers who do not want to be contacted')

        # copy dataframe keeping only relevant matching columns
        df_copy = df[all_col].copy()

        # clean data and convert data types
        with pd.option_context('mode.chained_assignment', None):
            # string columns
            logger.debug('Processing string columns')
            df_copy.loc[:, [*nme, *eml]] = df_copy[[*nme, *eml]].apply(lambda x: x.str.lower())
            df_copy.loc[:, nme] = df_copy[nme].replace(r"[^a-zA-Z]+", '', regex=True)

            # numeric columns
            logger.debug('Processing numeric columns')
            df_copy.loc[:, [*num, *fax]] = df_copy[[*num, *fax]].replace(r'[^0-9]+', '', regex=True)
            rgx_drop = '|'.join([f'^{pre}' for pre in self.phone_drop])
            df_copy.loc[:, num] = df_copy[num].replace(rgx_drop, '', regex=True)

            df_copy.loc[:, [*num, *fax]] = df_copy[[*num, *fax]].apply(pd.to_numeric, errors='coerce')

        # attempt to drop duplicates
        if 'd' in self.options and (kind != 'primary' or 'z' in self.options):
            df_copy = self.drop_dupes(df_copy, cols)

        return df_copy

    def get_df(self, filename, kind='primary', sep=';', nrows=20):
        """load and process dataframe from input folder"""
        text = 'primary' if kind == 'primary' else 'customer'

        # get number of header rows to skip
        skip = -1
        max_length = -1
        with open(f'{self.input_path}/{filename}', 'r') as file:
            for i, line in enumerate(file):
                if i > nrows:
                    break
                length = len([word for word in line.split(sep) if len(word) > 0])
                if length > max_length:
                    skip = i
                    max_length = length

        # load database
        logger.info(f'Loading {text} database... ')
        logger.debug(f'Separator: {sep}')
        logger.debug(f'Looking for header in first {nrows} rows')
        logger.debug(f'Skip to row: {skip}')
        time = perf_counter()
        df_main = pd.read_csv(f'{self.input_path}/{filename}', sep=sep,
                              skiprows=skip, low_memory=False).dropna(how='all')
        if kind != 'primary':
            df_main = df_main.reset_index().rename({'index': 'cust_index'}, axis=1)
        logger.info(f'Done: {round(perf_counter() - time, 5)}s\n')

        # clean dataframe
        logger.info(f'Cleaning {text} dataframe... ')
        time = perf_counter()
        df_copy = self.clean_df(df_main, kind)

        logger.debug(f'Main shape: {df_main.shape}  -  Copy shape: {df_copy.shape}')
        logger.info(f'Done: {round(perf_counter() - time, 5)}s\n')

        return df_main, df_copy

    def convert_excel(self):
        """convert excel files to csv for faster reading in pandas"""

        csv_files = glob.glob(f'{self.input_path}/*.csv')
        xl_files = glob.glob(f'{self.input_path}/*.xlsx')
        commands = []

        logger.debug(f'Found {len(csv_files)} csv files')
        logger.debug(f'Found {len(xl_files)} excel files to convert')

        # build list of functions for xlsx2csv.py
        for filepath in xl_files:
            filename = re.search(r'(.+[\\|/])(.+)(\.(csv|xlsx|xlx))', filepath)
            # if not already converted and not temp file
            if filename[0].replace('xlsx', 'csv') not in csv_files and not filename.group(2).startswith('~$'):
                call = ["./xlsx2csv.py", filepath, join(self.input_path, filename.group(2) + '.csv'), "-d", ";"]
                commands.append(call)

        if len(commands) == 0:  # return if no excel files or all converted
            self.move_excel(xl_files)
            return

        logger.info('Converting Excel files to csv... ')
        logger.debug(f'Built {len(commands)} commands')
        time = perf_counter()
        if sys.platform == "win32":
            with Pool(cpu_count()) as pool:  # parallelize and convert
                [pool.apply_async(xlsx2csv.run, args=call).get() for call in commands]
        else:
            [xlsx2csv.run(call) for call in commands]
        logger.info(f'Done: {round(perf_counter() - time, 5)}s')

        self.move_excel(xl_files)

    def move_excel(self, xl_files):
        if len(xl_files) == 0:
            return

        print()
        logger.info("Moving redundant Excel files to 'redundant' folder...")
        redundant = join(dirname(self.input_path), 'redundant')
        print(redundant)
        if not os.path.exists(redundant):  # if folder doesn't exist
            os.makedirs(redundant)  # create folder

        [shutil.move(filepath, join(redundant, basename(filepath))) for filepath in xl_files]
        logger.info('Done.\n')
        print()

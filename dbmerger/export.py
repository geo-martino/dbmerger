import json
import logging
import os
import re
import sys
from itertools import chain
from os.path import join
from time import perf_counter

import numpy as np
import pandas as pd
from openpyxl import load_workbook

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
stdout_handler = logging.StreamHandler(stream=sys.stdout)
stdout_handler.setLevel(logging.INFO)
logger.addHandler(stdout_handler)


class Export:

    def __init__(self, settings):

        # options
        self.options = settings['options']

        # column names
        self.prim_cols = settings['prim_cols']
        self.prim_id = settings['id']
        self.prim_interest = settings['interest']
        self.prim_greeting = settings['greeting']
        self.prim_title = settings['title']

        self.cust_cols = settings['cust_cols']
        self.cust_title = settings['cust_title']

        # mailing list options
        self.email_sort = settings['email_sort']
        self.email_add = settings['email_add']
        self.phone_drop = settings['phone_drop']
        self.phone_sort = settings['phone_sort']
        self.title_drop = settings['title_drop']

        # directories
        self.check_path = settings['check']
        self.final_path = settings['final']
        self.open = settings['open']

        # create directories as needed
        if not os.path.exists(self.check_path):
            logging.debug(f'Creating check path at {self.check_path}')
            os.makedirs(self.check_path)

        if not os.path.exists(self.final_path):
            logging.debug(f'Creating final path at {self.final_path}')
            os.makedirs(self.final_path)

        logger.debug(f'========== Export object instantiated ==========\n{json.dumps(self.__dict__, indent=2)}')

    def check(self, filename, same, diff, none):
        """Export 3 match conditions dataframes to excel sheet"""

        filename += '_check'
        filepath = join(self.check_path, f'{filename}.xlsx')

        print('-' * 88)
        logger.info('Exporting check spreadsheet...')
        logger.debug(f'Filename: {filename}.xlsx')
        logger.debug(f'Shapes: same {same.shape}, diff {diff.shape}, none {none.shape}')

        logger.debug(f'Saving check...')
        i = 0
        while True:
            try:
                with pd.ExcelWriter(filepath) as writer:  # export to excel
                    same.set_index(self.prim_id).to_excel(writer, sheet_name='Perfect Match')
                    diff.set_index(self.prim_id).to_excel(writer, sheet_name='Uncertain Match')
                    none.set_index(self.prim_id).to_excel(writer, sheet_name='Missing Data')
                break
            except PermissionError:  # prompt user to close file before continuing
                if 'q' not in self.options:
                    logger.error(f'ERROR: close the file {filename}.xlsx and hit enter to continue: ')
                    input()
                    print('\nRetrying...')
                else:
                    i += 1
                    filename += f'_{i}'
                    filepath = join(self.final_path, f'{filename}.xlsx')

        print(f"Match cases exported to: \n{filepath}")
        logger.debug(f"Match cases exported to: \n{filepath}\n")
        print('-' * 88, '\n')

    def mail(self, filename, matches, prim_main, cust_main, kind='perfect'):
        """Add mailing list sheet to final export"""

        filename += '_final'

        logger.info(f'Processing {kind.lower()} mailing list...')
        print()
        logger.debug(f'Filename: {filename}.xlsx')
        logger.debug(f'Shapes: matches {matches.shape}, primary {prim_main.shape}, cust {cust_main.shape}')
        filepath = join(self.final_path, f'{filename}.xlsx')

        # merge data
        logger.debug('Merging data')
        mail_list = matches[[self.prim_id, 'cust_index']].merge(prim_main, on=self.prim_id)
        mail_list = mail_list.merge(cust_main, on='cust_index', suffixes=(None, '_cust'))

        # filter columns
        logger.debug('Filtering columns')
        prim_cols = {k: list(set(v).intersection(set(mail_list.columns))) for k, v in self.prim_cols.copy().items()}
        cust_cols = {k: set(v).intersection(set(mail_list.columns)) for k, v in self.cust_cols.copy().items()}

        email_cols = list({*prim_cols['e'], *list(cust_cols['e'].intersection(set(cust_main.columns)))})
        phone_cols = list({*prim_cols['p'], *list(cust_cols['p'].intersection(set(cust_main.columns)))})

        all_cols = ['cust_index', self.prim_id, *self.email_add, *self.prim_cols['n'], *email_cols, *phone_cols]

        logger.debug(f'Filtered emails: {email_cols}')
        logger.debug(f'Filtered phones: {phone_cols}')
        logger.debug(f'Filtered all:\n {all_cols}\n')

        # filter columns and drop empty rows
        mail_list = mail_list[all_cols].dropna(how='all')

        # process and order column data
        mail_list = self.get_mailing_email(mail_list, email_cols)
        mail_list = self.get_mailing_phones(mail_list, phone_cols)
        mail_list = self.get_mailing_greetings(mail_list)
        mail_list = self.drop_mail_dupes(mail_list, email_cols, phone_cols)

        # drop empty columns
        mail_list = mail_list.replace(r'\s+( +\.)', np.nan, regex=True).replace('', np.nan)

        # split sheets for those with emails
        missing_list = mail_list[mail_list['Email 1'].isna()].dropna(how='all', axis=1)
        mail_list = mail_list[mail_list['Email 1'].notna()].dropna(how='all', axis=1)

        print('-' * 88, sep='')
        logger.info('Exporting mailing list...')
        i = 0
        while True:  # save
            try:
                book = load_workbook(filepath)  # load final matches spreadsheet
                with pd.ExcelWriter(filepath, engine='openpyxl') as writer:  # export to excel
                    writer.book = book  # add final matches
                    writer.sheets = {ws.title: ws for ws in book.worksheets}  # rename sheets
                    mail_list.sort_values(by=self.prim_id).to_excel(writer, index=False,
                                                                    sheet_name=f'{kind.title()} Mailing List')
                    missing_list.sort_values(by=self.prim_id).to_excel(writer, index=False,
                                                                       sheet_name=f'{kind.title()} Missing Emails List')
                break
            except PermissionError:  # prompt user to close file before continuing
                if 'q' not in self.options:
                    logger.error(f'ERROR: close the file {filename}.xlsx and hit enter to continue: ')
                    input()
                    print('\nRetrying...')
                else:
                    i += 1
                    filename += f'_{i}'
                    filepath = join(self.final_path, f'{filename}.xlsx')

        print(f"Added mailing list sheets to final matches export: \n{filepath}")
        logger.debug(f"Added mailing list sheets to final matches export: \n{filepath}\n")
        print('-' * 88, '\n')

    def final(self, filename, final_matches, prim_main, cust_main):
        """Create and save final excel sheet"""

        filename += '_final'

        print('-' * 88)
        logger.info('Exporting final matches...')
        logger.debug(f'Filename: {filename}.xlsx')
        logger.debug(f'Shapes: final_matches {final_matches.shape}, primary {prim_main.shape}, cust {cust_main.shape}')

        filepath = join(self.final_path, f'{filename}.xlsx')
        final_IDs = final_matches[[self.prim_id, 'cust_index']]

        # filter original dataframes
        logger.debug('Filtering original database')
        primary = prim_main[prim_main[self.prim_id].isin(final_IDs[self.prim_id])].sort_values(by=self.prim_id)
        prim_cols = primary.columns

        # add customer columns
        if any(char in self.options for char in ['c', 'm', 'a']):
            logger.debug('Adding customer data')

            # filter customer columns to only those used for merging
            if 'm' in self.options:
                logger.debug('Filtering customer database columns to only columns used for matching')
                cust_cols = [col for col in list(chain(*self.cust_cols.values())) if col in cust_main.columns]
                cust_main = cust_main[['cust_index', *cust_cols]]

            with pd.option_context('mode.chained_assignment', None):
                primary.loc[:, 'CUSTOMER DATA >>'] = 'CUSTOMER DATA >>'
            primary = primary.merge(final_IDs, on=self.prim_id).merge(cust_main, on='cust_index',
                                                                      suffixes=(None, '_cust')).dropna(how='all')

            # split columns if append to extra sheet option enabled
            if 'a' in self.options:
                logger.debug('Splitting columns to append to separate sheet')
                cust = primary[primary.columns[len(prim_cols) + 1:]]
                primary = primary[primary.columns[:len(prim_cols)]]

        logger.debug(f'Saving final...')
        i = 0
        while True:
            try:
                with pd.ExcelWriter(filepath) as writer:  # export to excel
                    primary.to_excel(writer, index=False, sheet_name='Primary Data')
                    if 'a' in self.options:  # append customer data to separate sheet
                        cust.to_excel(writer, index=False, sheet_name='Customer Data')
                break
            except PermissionError:  # prompt user to close file before continuing
                if 'q' not in self.options:
                    logger.error(f'ERROR: close the file {filename}.xlsx and hit enter to continue: ')
                    input()
                    print('\nRetrying...')
                else:
                    i += 1
                    filename += f'_{i}'
                    filepath = join(self.final_path, f'{filename}.xlsx')

        print(f"Final matches exported to: \n{filepath}")
        logger.debug(f"Final matches exported to: \n{filepath}\n")
        print('-' * 88, '\n')

    def drop_mail_dupes(self, mail_list, email_cols, phone_cols):
        """merge data from duplicate entries, with duplicates determined as those with matching cust_index"""
        logger.info('Merging duplicates... ')
        # log stats
        time = perf_counter()
        start_len = len(mail_list)

        # drop exact row duplicates
        mail_list.drop_duplicates(inplace=True)
        # get list of duplicated customer indices
        dupes = mail_list[mail_list.duplicated(subset='cust_index')]
        indices = set(dupes['cust_index'].values)
        logger.debug(f'{len(dupes)} duplicates found - {len(indices)} unique')

        def condense(i, cols, kind='phone'):
            """get condensed list of unique data for a customer index"""
            # list of all values that are not null
            values = set([x for x in mail_list[mail_list['cust_index'] == i][cols].values.flatten() if len(x) > 0])
            if kind == 'phone':
                clean_numbers = {re.sub(r"[^0-9]+", '', num): num for num in values
                                 if len(re.sub(r"[^0-9]+", '', num)) > 0}
                return list(clean_numbers.values())
            else:
                return [email.strip() for email in values if len(email.strip()) > 1]

        def unique_titles(i):  # get condensed string of unique titles
            """get unique titles from from a given dataframe index"""
            # convert values to string
            titles = ' '.join(mail_list[mail_list['cust_index'] == i][self.prim_title].values.flatten())
            # split values
            titles = [i.strip() for i in titles.split('.')]
            # get unique titles
            unique = set(title.lower() for title in titles if len(title) > 1)
            reduced = []
            for title in titles:
                if title.lower() in unique:  # if title is unique
                    reduced.append(title + '.')
                    unique.remove(title.lower())  # remove form unique list
            return ' '.join(reduced)  # return string of titles

        # merge unique data for each data
        logger.debug('Merging data')
        emails = {idx: condense(idx, email_cols, 'email') for idx in indices}
        phones = {idx: condense(idx, phone_cols, 'phone') for idx in indices}
        titles = {idx: unique_titles(idx) for idx in indices}

        logger.debug('Creating dataframes')
        # rename columns and create dataframes for each collection of merged data
        new_cols = [f'Email {i + 1}' for i in range(max(len(v) for i, v in enumerate(emails.values())))]
        emails_df = pd.DataFrame.from_dict(emails, orient='index', columns=new_cols)
        new_cols = [f'Phone Number {i + 1}' for i in range(max(len(v) for i, v in enumerate(phones.values())))]
        phones_df = pd.DataFrame.from_dict(phones, orient='index', columns=new_cols)
        titles_df = pd.DataFrame.from_dict(titles, orient='index', columns=[self.prim_title])

        # create dataframe of all merged data
        all_df = pd.concat([emails_df, phones_df, titles_df], axis=1)

        # extract non-merged columns for merged data and join merged
        dupes = dupes[set(dupes.columns) - {*email_cols, *phone_cols, self.prim_title}].set_index('cust_index')
        dupes = dupes.join(all_df).reset_index().rename(columns={'index': 'cust_index'})
        dupes.drop_duplicates(subset='cust_index', inplace=True)

        logger.debug('Renaming contact data columns')
        # rename mail_list columns to match dupe data
        mail_list.rename(columns={col: f'Email {i + 1}' for i, col in enumerate(email_cols)}, inplace=True)
        mail_list.rename(columns={col: f'Phone Number {i + 1}' for i, col in enumerate(phone_cols)}, inplace=True)

        # drop all duplicate indices
        mail_list.drop_duplicates(subset='cust_index', keep=False, inplace=True)

        # concatenate both frames
        logger.debug('Concatenating dataframes')
        logger.debug(f'Dropped {start_len - len(mail_list)} duplicates')
        logger.debug(f'Start: {start_len}, End: {len(mail_list)}')
        logger.info(f'Done: {round(perf_counter() - time, 5)}s\n')
        return pd.concat([mail_list, dupes], axis=0, ignore_index=True)

    def get_mailing_email(self, mail_list, cols):
        """sort and drop duplicate emails per index in mailing list"""
        if len(cols) <= 1:
            logger.debug('Skipping sort emails: cols length <= 1')
            return mail_list

        # enumerate order of items in email drop list to dict
        email_order = {c: p + 1 for (p, c) in enumerate(self.email_sort)}
        # missing emails placed at end of order
        email_order['@'] = len(self.email_sort) + 1

        def sort_emails(emails):
            # returns sorted list of emails according to above order
            # emails not in order dict i.e. good emails, placed at position 0
            max_len = len(emails)
            emails = set([x.lower().replace(' ', '') for x in emails if '@' in x])
            sorted_emails = sorted(emails, key=lambda c: email_order.get(c.split('@')[0] + '@', 0))
            return pd.Series(sorted_emails + [''] * (max_len - len(sorted_emails)))

        logger.info('Sorting emails... ')
        logger.debug(f'Email Order: {email_order}')

        try:  # apply sort per row of numpy array
            time = perf_counter()
            mail_list[cols] = mail_list[cols].fillna('').apply(lambda x: sort_emails(x), axis=1).to_numpy()
            logger.info(f'Done: {round(perf_counter() - time, 5)}s\n')
        except ValueError:
            logger.exception('Error: Skipping email sorting (ValueError: df is empty)\n')

        return mail_list

    def get_mailing_phones(self, mail_list, cols):
        """sort and drop duplicate phone numbers per index in mailing list"""
        if len(cols) <= 1:
            logger.debug('Skipping sort phones: cols length = 1')
            return mail_list

        # enumerate order of items in phone drop list to dict
        phone_order = {c: p + 1 for (p, c) in enumerate(self.phone_sort)}

        # define regex conditions based on sort and drop
        rgx_sort = '|'.join([f'^{pre}' for pre in self.phone_sort])
        rgx_drop = '|'.join([f'^{pre}' for pre in self.phone_drop])

        def sort_phones(phones):
            """returns sorted list of phone numbers according to above order,
            numbers not in order dict placed at penultimate position"""

            # drop duplicates
            max_len = len(phones)
            phones = set(phones)

            # clean: keep only digits in numbers
            clean_raw = {re.sub(r'[^0-9]+', '', num): num for num in phones if len(num) > 0}

            # drop duplicate numbers if missing prefixes
            for num in list(clean_raw.keys()):
                num_trunc = re.sub(rgx_drop, "", num)
                if num_trunc != num and num_trunc in clean_raw.keys():
                    del clean_raw[num_trunc]

            # extract prefixes from raw numbers
            raw_prefixes = {num: num_clean.replace(re.sub(rgx_sort, "", num_clean), '')
                            for num_clean, num in clean_raw.items()}

            def get_order(number):
                """get the position of a given number"""
                if len(number) == 0:
                    return len(self.phone_sort) + 2
                else:
                    return phone_order.get(raw_prefixes.get(number, 'none'), len(self.phone_sort) + 1)

            # sort phone numbers
            sorted_phones = sorted(clean_raw.values(), key=lambda c: get_order(c))

            # add extra empty strings to match length for numpy function to work
            return pd.Series(sorted_phones + [''] * (max_len - len(sorted_phones)))

        logger.info('Sorting phone numbers... ')
        logger.debug(f'Phone Order: {phone_order}')
        logger.debug(f'Removing titles with regex: {rgx_drop}')

        try:  # apply sort per row of numpy array
            time = perf_counter()
            mail_list[cols] = mail_list[cols].fillna('').apply(lambda x: sort_phones(x), axis=1).to_numpy()
            logger.info(f'Done: {round(perf_counter() - time, 5)}s\n')
        except ValueError:
            logger.exception('Error: Skipping phone sorting (ValueError: array is empty)\n')

        return mail_list

    def get_mailing_greetings(self, mail_list):
        """clean, sort, and drop duplicates for greetings and titles"""
        def sort_titles(titles):
            """get condensed string of unique titles"""
            # split values, dropping the defined drop titles
            titles = [i.strip() for i in titles.split('.') if len(re.sub(r"[^a-z]", '', i.lower())) > 1
                      and re.sub(r"[^a-z]", '', i.lower()) not in self.title_drop]
            # capitalise titles
            titles = [i.title() if (i == i.lower() and 'pd' not in i.lower()) else i for i in titles]
            titles = [i + '. med' if 'dipl' in i.lower() else i for i in titles]
            # get unique titles
            unique = set(title.lower() for title in titles if len(title) > 1)
            reduced = []
            for title in titles:
                if title.lower() in unique:  # if title is unique
                    reduced.append(title + '.')
                    unique.remove(title.lower())  # remove form unique list
            return ' '.join(reduced)  # return string of titles

        logger.info('Sorting greetings and titles... ')
        time = perf_counter()
        # fix typos
        if self.prim_greeting in mail_list.columns:
            mail_list[self.prim_greeting] = mail_list[self.prim_greeting].str.replace('herte', 'hrte').fillna(
                'Sehr geehrte/r Frau/Herr')

        # clean titles
        if self.prim_title in mail_list.columns:
            rgx_drop = '|'.join([f'(?i){pre}' for pre in [*self.title_drop, 'nan']])
            logger.debug(f'Removing titles with regex: {rgx_drop}')
            # drop titles
            prim_titles = mail_list[self.prim_title].str.replace(rgx_drop, '', regex=True).fillna('')

            try:
                # get customer title column
                cust_title = list(set(self.cust_title).intersection(mail_list.columns))[0]
                # drop titles
                cust_titles = mail_list[cust_title].str.replace(rgx_drop, '', regex=True).fillna('')
                # merge titles
                mail_list[self.prim_title] = (prim_titles + ' ' + cust_titles).apply(lambda x: sort_titles(x))
            except KeyError:
                logger.error('Error: Title column not found in customer database, using only primary titles... ')
                mail_list[self.prim_title] = prim_titles.apply(lambda x: sort_titles(x))

        logger.info(f'Done: {round(perf_counter() - time, 5)}s\n')
        return mail_list

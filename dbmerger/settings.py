import json
import logging
import os
import re
import shutil
import sys
from inspect import cleandoc
from itertools import chain
from os.path import expanduser, join

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
stdout_handler = logging.StreamHandler(stream=sys.stdout)
stdout_handler.setLevel(logging.INFO)
logger.addHandler(stdout_handler)


class Settings:
    """Get user input to set custom settings and columns names"""

    def __init__(self):
        # if settings file doesn't exist, restore from backup
        if not os.path.isfile('settings.txt'):
            logger.warning('Settings file not found. Restoring from default.')
            print()
            shutil.copyfile('settings.bak', 'settings.txt')

        # system specific settings/commands
        if sys.platform == "linux":
            self.open = "xdg-open "
            self.folder = join(join(expanduser('~')), 'Desktop')
        elif sys.platform == "darwin":
            self.open = "open "
            self.folder = join(join(expanduser('~')), 'Desktop')
        elif sys.platform == "win32":
            self.open = ""
            self.folder = join(join(os.environ['USERPROFILE']), 'Desktop')

        # check for errors in settings file
        while True:
            try:
                self.all_settings = self.load_settings()
                break
            except KeyError:
                logger.error("ERROR: Settings file keys not recognised. Inspect settings.txt to fix.")
                print("Would you like to delete settings.txt and restore from default settings backup?")
                change = 'y' in input("(type 'y' (yes) or 'n' (no)) : ").lower().strip()
                print()
                if change:
                    shutil.copyfile('settings.bak', 'settings.txt')
                    print('-' * 88, '\n')

        # system specific open command
        self.all_settings['open'] = self.open

        # assign settings to self to reduce coding clutter
        self.options = self.all_settings['options']
        if 'n' not in self.options:
            self.options += ['n']
        self.prim_cols = self.all_settings['prim_cols']
        self.cust_cols = self.all_settings['cust_cols']

        # get column type keys
        self.keys = list(self.prim_cols.keys())

        # initial default options print
        self.print_options(True, 's' in self.options)

        logger.debug(f'========== Settings object instantiated ==========\n{json.dumps(self.__dict__, indent=2)}')

    def load_settings(self):
        """load settings from settings.txt file in program directory"""
        logger.debug('Loading settings from settings.txt')

        strings = []  # store settings strings
        with open('settings.txt', 'r') as file:
            for line in file:  # check if line is options line
                if line[0] != '#' and len(line.strip().split(':')) > 1:
                    line = line.split(':')
                    # only split for first ':'
                    line = [line[0], ' '.join(line[1:])]
                    strings.append(line)

        # general settings clean and split to dict
        all_settings = {k.lower().strip().replace(' ', '_'): s.strip() for k, s in strings}
        logger.debug(f'Raw settings\n{json.dumps(all_settings, indent=2)}')

        # options
        options = list(set(all_settings['options'].lower().replace("'", "").replace(" ", "")))
        if 'q' in options:
            options.append('s')
        if {'c', 'm'}.issubset(set(options)):
            options.remove('m')

        # column names
        id = all_settings['id']
        interest = all_settings['interest']
        greeting = all_settings['greeting']
        title = all_settings['title']
        prim_cols = {k[4]: [v.strip() for v in s.split(',')]
                     for k, s in all_settings.items() if k[0:3] == 'pri' and k != 'primary_data'}

        cust_cols = {k[4]: [v.strip() for v in s.split(',')]
                     for k, s in all_settings.items() if k[0:3] == 'cus'}
        cust_cols = {k: v + [f'{i}_cust' for i in v] if k != 'n' else v for k, v in cust_cols.items()}
        cust_title = cust_cols['t']
        cust_areas = cust_cols['a']
        del cust_cols['t']
        del cust_cols['a']

        # mailing list/phone options
        email_sort = [v.strip() for v in all_settings['email_sort'].split(',')]
        email_add = [v.strip() for v in all_settings['email_add'].split(',')]
        phone_sort = [v.strip() for v in all_settings['phone_sort'].split(',')]
        phone_drop = [v.strip() for v in all_settings['phone_drop'].split(',')]
        title_drop = [v.strip().lower() for v in all_settings['title_drop'].split(',')]

        # directories
        folder = join(self.folder, re.sub(r'[\\/*?:"<>|]', "", all_settings['folder_name']))
        paths = {k: join(folder, v.replace("'", '')).replace(' ', '_')
                 for k, v in all_settings.items() if k in ['input', 'check', 'final']}
        prim_filename = re.sub(r'[\\/*?:"<>|]', '', all_settings['primary_data'])

        # load field column translations
        strings = []  # store strings
        with open('areas.txt', 'r') as file:
            for line in file:  # check if line is options line
                if line[0] != '#' and len(line.strip().split(':')) > 1:
                    line = line.split(':')
                    # only split for first ':'
                    line = [line[0], ' '.join(line[1:])]
                    strings.append(line)

        # clean specialities and split to dict
        areas = {k.strip(): s.strip().split(',') for k, s in strings}
        logger.debug(f'Loaded field column translations')

        clean_settings = {
            'options': options,
            'prim_cols': prim_cols,
            'id': id,
            'interest': interest,
            'greeting': greeting,
            'title': title,
            'areas': areas,
            'cust_cols': cust_cols,
            'cust_title': cust_title,
            'cust_areas': cust_areas,
            'email_sort': email_sort,
            'email_add': email_add,
            'phone_sort': phone_sort,
            'phone_drop': phone_drop,
            'title_drop': title_drop,
            'folder': folder.replace(' ', '_'),
            'input': paths['input'],
            'check': paths['check'],
            'final': paths['final'],
            'prim_filename': prim_filename,

        }

        logger.debug(f'Clean settings\n{json.dumps(clean_settings, indent=2)}')

        return clean_settings

    def print_options(self, initial=False, skip=False):
        """display current settings to user"""

        # define option strings for printing
        match_dict = {'n': '' if any(char in self.options for char in ['p', 'e', 'f']) else ' only',
                      'p': 'phone numbers', 'e': 'emails', 'f': 'faxes'}
        match_options = ', '.join([v for k, v in match_dict.items() if k in self.options])
        drop_option = "Attempt" if 'd' in self.options else "Do not attempt"

        export_option = "Export" if 'x' in self.options else "Do not export"
        mailing_option = "Create" if 'l' in self.options else "Do not create"

        cust_option = "Add matching" if 'm' in self.options else "Do not include any"
        cust_option = "Add all" if 'c' in self.options else cust_option
        append_option = "a separate sheet\n" if 'a' in self.options else "primary sheet\n"
        append_option = f"- Add customer data to {append_option}" \
            if any(char in self.options for char in ['c', 'm']) else ""

        if initial:  # print for initial load
            print(cleandoc(f"""\
            The loaded options for this program are as follows:

            - Files loaded from {self.all_settings['folder']}
            - Match names{match_options}
            - {drop_option} to drop duplicates from primary database
            - {export_option} intermediary match cases to excel spreadsheets
            - {mailing_option} separate mailing list sheet
            - {cust_option} customer's data in final export
            {append_option}"""))
            if not skip:
                print('\n', '-' * 88, '\n', sep='')
                print(cleandoc("""
                To change these options, please input one or many of the following characters:
                (leave blank to use the above options)
                
                (note: if certain types of options are not specified then default options will be used)
                (note: match conditions may be combined i.e. typing 'ef' matches by name, 
                confirming with email and fax)

                # MATCH OPTIONS
                - n: match by name only (enabled by default in the program)
                - p: match by phone number
                - e: match by email
                - f: match by fax
                - d: attempt to drop duplicate entries from customer database based on phone,
                     email, and fax number columns
                - z: if 'd' enabled, also attempt to drop duplicates from primary database
                - i: drop customer's who do not want to be contacted (i.e., where Interesse MAFO = 2)

                # EXPORT OPTIONS
                - x: export intermediary match cases to Excel spreadsheets
                - l: append mailing list sheet
                - c: append all customer's data in final export to primary sheet
                - m: append matching customer's data to the primary sheet in final export to primary sheet
                - a: add customer data as an extra sheet rather than to the primary sheet
                - u: add extra mailing sheet for uncertain matches

                ### RUNTIME OPTIONS
                - s: fast mode - skip all settings checks and just use the settings in this file
                - v: view mode - view conflicting entries in program and resolve by manual ID input
                                 instead of by external Excel sheet.
                - q: auto mode - skip all user input and automatically run the program 
                                 (enables fast mode by default)
                - o: omit mode - when auto mode is enabled, only keep perfect matches
                """), end='\n\n')
        else:  # print current options
            print(cleandoc(f"""\n{'-' * 88}
            {'=' * 56}
            Options selected:
            - Files loaded from {self.all_settings['folder']}
            - Match names{match_options}
            - {export_option} intermediary match cases to excel spreadsheets
            - {cust_option} customer's data in final export
            - {mailing_option} separate mailing list sheet
            {'=' * 56}"""))

    def get_options(self):
        """Get user input for custom options"""
        logger.debug('Getting user input for options')

        print('-' * 88)
        print('Choose options (leave blank to use default options)')
        options = input("input as one string of characters (e.g. 'pec' or 'ml') : ").lower().strip()
        options = list(set(options.replace("'", "").replace(" ", "")))
        logger.debug(f'input: {options}')

        # enable fast mode if auto mode
        if 'q' in options:
            options.append('s')

        # check for conflicting customer data options, prefer 'keep all'
        if {'c', 'm'}.issubset(set(options)):
            logger.debug('Dropping m option')
            options.remove('m')

        if len(options) != 0:  # add names option if not given
            self.all_settings['options'] = list(set(options + ['n']))

        logger.debug(f"output: {self.all_settings['options']}")
        self.print_options()

    def print_columns(self):
        """print current column name settings"""

        areas = len(self.all_settings['areas'].keys())
        alt = len(list(chain(*self.all_settings['areas'].values())))
        logger.debug(f'{areas} Columns Defined - {alt} Alternate Names Given')

        print(cleandoc(f"""{'-' * 88}
        
        Primary database default column names to search for matches are as follows:

        - Names: {self.prim_cols['n']}
        - Phone: {self.prim_cols['p']}
        - Email: {self.prim_cols['e']}
        - Fax  : {self.prim_cols['f']}
        
        - ID       : {self.all_settings['id']}
        - Interest : {self.all_settings['interest']}
        - Greeting : {self.all_settings['greeting']}
        - Title     : {self.all_settings['title']}
        - Areas     : {areas} Columns Defined ({alt} Alternate Names Given)

        Customer database default column names to search for matches are as follows:

        - Names: {self.cust_cols['n']}
        - Phone: {self.cust_cols['p']}
        - Email: {self.cust_cols['e']}
        - Fax  : {self.cust_cols['f']}
        
        - Title:  {self.all_settings['cust_title']}
        - Areas:  {self.all_settings['cust_areas']}"""), end='\n\n')

        print('-' * 88)

    def change_column_names(self):
        """Ask user if they want to change default column names and run methods"""

        # print default column names
        self.print_columns()

        # ask user to change column names
        print("Would you like to change any of these column names?")
        change = 'y' in input("(type 'y' (yes) or 'n' (no)) : ").lower().strip()

        if change:  # methods for changing column names
            self.get_column_names('primary')
            self.get_column_names('cust')

        print('-' * 88, '\n')

    def get_column_names(self, kind='primary'):
        """Change default column names"""

        name = 'Primary' if kind == 'primary' else 'Customer'
        cols = self.prim_cols if kind == 'primary' else self.cust_cols
        logger.debug(f'Getting user input for {name} columns')

        print('-' * 88)
        print(f'\n---{name} Column Names---\n')

        while True:
            print(f'Which column types would you like to change from the {name} database?')
            print('(leave blank for default names)')

            # get input from user for which column names to change
            col_change = list(input('n = names, p = phone numbers, e = emails, f = fax numbers : ').lower().strip())

            # check user has defined column names to change
            if any(char in col_change for char in self.options):
                print('\n' + '-' * 88 + '\n')
                print('Type column names as comma-separated list (case-sensitive)')
                print('(e.g. Telefon-Festnetz, Telefon-Mobil (beruflich), Telefon-Mobil (privat))\n')

                if 'n' in col_change:  # name columns
                    inp = input('- Define name columns: ')
                    cols['n'] = [x.strip() for x in inp.split(',')] if inp != '' else cols['n']

                if 'p' in col_change and 'p' in self.options:  # phone number columns
                    inp = input('- Define phone number columns: ')
                    cols['p'] = [x.strip() for x in inp.split(',')] if inp != '' else cols['p']

                if 'e' in col_change and 'e' in self.options:  # email columns
                    inp = input('- Define email columns: ')
                    cols['e'] = [x.strip() for x in inp.split(',')] if inp != '' else cols['e']

                if 'f' in col_change and 'f' in self.options:  # fax number columns
                    inp = input('- Define fax number columns: ')
                    cols['f'] = [x.strip() for x in inp.split(',')] if inp != '' else cols['f']
                break
            elif all(char in self.keys for char in col_change):  # columns types not in options
                break
            elif not col_change:  # use default columns
                break
            else:  # not recognised, loop back
                print()
                logger.error('!!! ERROR: Column types not recognised !!!')
                print()

        # print column names to use
        print(cleandoc(f"""\n{'-' * 88}
        {'=' * 56}
        "Using the following column names for {name} database:"

        - Names: {cols['n']}
        - Phone: {cols['p']}
        - Email: {cols['e']}
        - Fax  : {cols['f']}
        {'=' * 56}"""))

        if kind == 'primary':
            self.prim_cols = cols
        else:
            self.cust_cols = {k: v + [f'{i}_cust' for i in v] if k != 'n' else v for k, v in cols.items()}

    def get_all(self):
        """run settings methods as appropriate"""
        if 's' not in self.options:  # if fast mode disabled
            self.get_options()

        if 's' not in self.options:  # if fast mode disabled
            self.change_column_names()
        else:
            logger.debug('FastMode: Skipping settings.get_all()')
            print('- Fast mode enabled')
            if 'v' in self.options:
                print('- Viewing and resolving conflicts in the console window')
            else:
                print('- Resolving conflicts through excel spreadsheet')
            print()
            self.print_columns()
            print()

        return self.all_settings

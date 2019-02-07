#!/usr/bin/python3
import sys, os
from pprint import pprint
from collections import defaultdict
import pandas as pd
import numpy as np


_outlook_fields = [ "First Name", "Middle Name", "Last Name", "Title", "Suffix",
    "Initials", "Web Page", "Gender", "Birthday", "Anniversary", "Location",
    "Language", "Internet Free Busy", "Notes", "E-mail Address", "E-mail 2 Address", "E-mail 3 Address",
    "Primary Phone", "Home Phone", "Home Phone 2", "Mobile Phone", "Pager", "Home Fax",
    "Home Address", "Home Street", "Home Street 2", "Home Street 3", "Home Address PO Box", "Home City",
    "Home State", "Home Postal Code", "Home Country", "Spouse", "Children", "Manager's Name",
    "Assistant's Name", "Referred By", "Company Main Phone", "Business Phone", "Business Phone 2", "Business Fax",
    "Assistant's Phone", "Company", "Job Title", "Department", "Office Location", "Organizational ID Number",
    "Profession", "Account", "Business Address", "Business Street", "Business Street 2", "Business Street 3",
    "Business Address PO Box", "Business City", "Business State", "Business Postal Code", "Business Country", "Other Phone",
    "Other Fax", "Other Address", "Other Street", "Other Street 2", "Other Street 3", "Other Address PO Box",
    "Other City", "Other State", "Other Postal Code", "Other Country", "Callback", "Car Phone",
    "ISDN", "Radio Phone", "TTY/TDD Phone", "Telex", "User 1", "User 2",
    "User 3", "User 4", "Keywords", "Mileage", "Hobby", "Billing Information",
    "Directory Server", "Sensitivity", "Priority", "Private", "Categories",
    ]

_google_fields = [
    'Name', 'Given Name', 'Additional Name', 'Family Name', 'Yomi Name', 'Given Name Yomi',
    'Additional Name Yomi', 'Family Name Yomi', 'Name Prefix', 'Name Suffix', 'Initials', 'Nickname',
    'Short Name', 'Maiden Name', 'Birthday', 'Gender', 'Location', 'Billing Information',
    'Directory Server', 'Mileage', 'Occupation', 'Hobby', 'Sensitivity', 'Priority',
    'Subject', 'Notes', 'Language', 'Photo', 'Group Membership', 'E-mail 1 - Type',
    'E-mail 1 - Value', 'E-mail 2 - Type', 'E-mail 2 - Value', 'Phone 1 - Type', 'Phone 1 - Value', 'Phone 2 - Type',
    'Phone 2 - Value', 'Website 1 - Type', 'Website 1 - Value',
]


class Converter(object):

    #
    # Google to Ootlokook converter
    # @Todo: convert from ootlook to google, select source/target format from command line
    # @todo: add other format
    #

    ## Name
    #join(Name Prefix, Additional Name , Name, Given Name, Family Name ) ->  Middle Name
    #
    ## Phone
    #Phone 1 - Value-> Home Phone
    #Phone 2 - Value-> home phone 2
    #
    #
    ## Mail
    #E-mail 1 - value -> E-mail Address
    #E-mail 2 - Value -> E-mail 2 Adress
    #
    ## Website
    #Website 1 - Value -> web page
    #Website 1 - Value -> web page 2
    #
    ## Group
    #Group Membership (remove '*' ':::' separated) -> Categories ';' separated

    def __init__(self, settings=None):
        self.s = settings

        self.fields = ['name', 'phone', 'mail', 'website', 'group']

        self._inner_sep = ':::'
        self.rows_in = _google_fields
        self.rows_out = _outlook_fields


    # no effect...
    def convert_wrapper(fun):
        def wrap(self, x):
            res = fun(self, x)
            return res
        return wrap

    @convert_wrapper
    def convert_name(self, x):
        res = {}
        #name = ' '.join([str(x.get(f, '')).strip() for f in ['Name Prefix', 'Additional Name' , 'Name', 'Given Name', 'Family Name']])
        name = x['Name']
        res['Middle Name'] = name
        return res

    @convert_wrapper
    def convert_phone(self, x):
        res = {}
        i = 1
        j = 1
        f = 'Phone %d - Value' % i
        f_res = 'Home Phone'
        while not pd.isna(x.get(f)):
            objs = str(x[f]).strip()
            for o in objs.split(self._inner_sep):
                res[f_res] = o
                f_res = 'Home Phone %d' % (i+j)
                j += 1
            i += 1
            f = 'Phone %d - Value' % (i)

        return res

    @convert_wrapper
    def convert_mail(self, x):
        res = {}
        i = 1
        j = 1
        f = 'E-mail %d - Value' % i
        f_res = 'E-mail Address'
        while not pd.isna(x.get(f)):
            objs = str(x[f]).strip()
            for o in objs.split(self._inner_sep):
                res[f_res] = o
                f_res = 'E-mail %d Address' % (i+j)
                j += 1
            i += 1
            f = 'E-mail %d - Value' % (i)

        return res

    @convert_wrapper
    def convert_website(self, x):
        res = {}
        i = 1
        j = 1
        f = 'Website %d - Value' % (i)
        f_res = 'Web Page'
        while not pd.isna(x.get(f)):
            objs = str(x[f]).strip()
            for o in objs.split(self._inner_sep):
                res[f_res] = o
                f_res = 'Web Page %d' % (i+j)
                j += 1
            i += 1
            f = 'Website %d - Value' % (i)

        return res

    @convert_wrapper
    def convert_group(self, x):
        groups = x.get('Group Membership')
        if pd.isna(groups):
            return None
        groups = groups.split(':::')
        groups = list(map(lambda x: x.strip('*, '), groups))
        groups = ';'.join(groups)
        res = {'Categories' : groups}
        return res

    def process_row(self, row):
        res = {}
        for f in self.fields:
            try:
                conv = getattr(self, 'convert_'+f)
            except AttributeError as e:
                raise e('converter method doesnt exist for field: %s' % f)

            res.update(conv(row))

        return res




if __name__ == '__main__':

    fn = sys.argv[1]
    if not os.path.isfile(fn):
        print('not a file: %s' % fn)
        exit(2)

    contacts_source = pd.read_csv(fn)
    google_fields = contacts_source.columns
    print("number of contacts: %d" % len(contacts_source))

    conv = Converter()

    contacts_target = []
    for i, v in contacts_source.iterrows():
        res = conv.process_row(v)
        contacts_target.append(res)

    d = pd.DataFrame(contacts_target, columns=_outlook_fields)

    fn_out = ''.join(fn.split('.')[:-1]) + '_g2o.csv'
    print('writing to: %s' % fn_out)
    with open(fn_out, 'w') as _f:
        _f.write(d.to_csv(header=_outlook_fields, index=False))





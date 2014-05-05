#!/usr/bin/env python3

'''
This module converts VCF 2.01 (supporse from Android phone) to Outlook 2003 CSV

.. note::
    At the moment, extracted only formatted name and phone numbers.

'''

#------------------------------------------------------------------------------
# Imports
#------------------------------------------------------------------------------

import io
import re
import quopri
import csv
from operator import itemgetter

#------------------------------------------------------------------------------
# Metadata
#------------------------------------------------------------------------------

__version__ = '0.1.0.0b'

#------------------------------------------------------------------------------
# Convertor
#------------------------------------------------------------------------------


# CSV relaited functions & structures
# ---------------------------------------------------------------------

class CsvCollection:
    '''
    Collection of :class:`CsvStructure` instances.
    '''
    __slot__ = ['__items']

    __items = None
    
    def __init__(self):
        self.__items = []
    
    def append(self, item):
        if isinstance(item, CsvStructure):
            self.__items.append(item)
        else:
            raise TypeError((
                'Expected CsvStructure instance, given {wrong_item}.'
            ).format(
                wrong_item=type(item)
            )
    
    def extend(self, items):
        for item in items:
            self.append(item)

    def to_file(file_path):
        with open(file_path, 'w') as csv_file:
            csv_writer = csv.writer(csv_file)
            for item in self.__items:
                csv_writer.writerow(item.as_tuple())

    def __getattr__(self, attr):
        return getattr(self.__items, attr)

    def __getitem__(self, key):
        return self.__items[key]


class CsvStructure:
    '''
    This class implements fields structure of Outlook CSV file.
    '''

    __slots__ = ['__fields']

    #: Fields available in Outlook 2003 CSV file.
    __ordered_keys = (
        "Title","First Name","Middle Name","Last Name","Suffix",
        "Given Name Yomi","Family Name Yomi","Home Street","Home City",
        "Home State","Home Postal Code","Home Country","Company","Department",
        "Job Title","Office Location","Business Street","Business City",
        "Business State","Business Postal Code","Business Country",
        "Other Street","Other City","Other State","Other Postal Code",
        "Other Country","Assistant's Phone","Business Fax","Business Phone",
        "Business Phone 2","Callback","Car Phone","Company Main Phone",
        "Home Fax","Home Phone","Home Phone 2","ISDN","Mobile Phone",
        "Other Fax","Other Phone","Pager","Primary Phone","Radio Phone",
        "TTY/TDD Phone","Telex","Anniversary","Birthday","E-mail Address",
        "E-mail Type","E-mail 2 Address","E-mail 2 Type","E-mail 3 Address",
        "E-mail 3 Type","Notes","Spouse","Web Page"
    )

    #: Existed phone fields in CSV structure
    __phone_cells = (
        "Mobile Phone", "Primary Phone", "Other Phone", "Home Phone", 
        "Home Phone 2", "Business Phone", "Business Phone 2", 
        "Radio Phone", "TTY/TDD Phone", "Car Phone", "Company Main Phone", 
        "Assistant's Phone"
    )

    #: Storage of CSV data
    __fields = dict.fromkeys(__ordered_keys)

    __getter_by_keys = itemgetter(*__ordered_keys)

    def as_tuple(self):
        return self.__getter_by_keys(self.__fields)

    def __setitem__(self, key, value):
        if key in self.__ordered_keys:
            self.__fields[key] = value
        else:
            raise KeyError((
                'Key "{wrong_key}" not present in CSV structure, '
                'available keys:\n{available_keys}'
            ).format(
                wrong_key=key,
                available_keys='\n'.join(self.__ordered_keys)
            )

    def __getitem__(self, key):
        return self.__fields[key]
  
    # Name of contact
    # -----------------------------------------------------------------
    @property
    def name(self):
        return self.__fields['First Name']

    @name.setter
    def name(self, name):
        self.__fields['First Name'] = name
    
    # Phones of contact
    # -----------------------------------------------------------------
    @property
    def phones(self):
        pass

    @phones.setter
    def phones(self, phones_list):
        for offset, phone_number in enumerate(phones_list, start=1):
            cell_name = (self.__phone_cells[:offset][-1:] or (None,))[0]
            if cell_name is not None:
                self[cell_name] = phone_number


# VCF relaited functions
# ---------------------------------------------------------------------

rgx_fn = re.compile(r'FN[^:]*:(.+?)\n[\w;,=]+:', re.S)
def extract_formatted_name(vcard_item):
    ''' Formatted name extractor from vCard. '''
    return rgx_fn.search(vcard_item).group(1).replace('\n', '')


rgx_phone = re.compile(r'TEL[^:]*:(.+?)\n')
def extract_phones(vcard_item):
    ''' Phone extractor from vCard. '''
    return rgx_phone.findall(vcard_item)


def vcards_reader(file_object):
    ''' vCards reader by blocks.
    
    :param file_object: opened VCF file
    :type file_object: file_like_object

    :rtype: generator

    '''
    buffer = io.StringIO()
    is_vcard = False

    for line in file_object:
        buffer.write(line)
        if line.startswith('BEGIN:VCARD'):
            is_vcard = True
        elif line.startswith('END:VCARD'):
            is_vcard = False
            buffer.seek(0)
            yield buffer.read()
            buffer.truncate()


def vcf_to_csv(vcf_filepath, csv_filepath=None)
    csv_filepath = csv_filepath or (os.path.splittext(vcf_filepath)[0] + '.csv')
    csv_collection = CsvCollection()

    with open(vcf_filepath, 'r') as vcf_file:
        for vcard in vcards_reader(vcf_file):
            fn = extract_formatted_name(vcard)
            phones = extract_phones(vcard)
            
            csv_item = CsvStructure()
            csv_item.name = fn
            csv_item.phones = phones
            csv_collection.append(csv_item)

    csv_collection.to_file(csv_filepath)


if __name__ == '__main__':
    pass

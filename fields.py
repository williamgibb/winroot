# XXX Fill out docstring!
"""

"""
from __future__ import print_function
import logging
import re

from errors import FieldsError

log = logging.getLogger(__name__)
__author__ = 'wgibb'

ROOT_BIRTH = 'BIRTH'
ROOT_FINAL = 'FINAL'
# XXX Move these into a class?
digit_start = r'^[0-9]'
valid_python_identifer = r'^[a-zA-Z_]+[a-zA-Z0-9_]*$'
invalid_python_identifiers = r'[^a-zA-Z0-9_]'
replacement_str = 'X'


class IdentityFields(object):
    identity_fields = ['RootName',
                       'Location#',
                       'BirthSession',
                       'Tube#']

    identity_attributes = {}
    for k in identity_fields:
        new_key = str(k)
        if not re.search(valid_python_identifer, new_key):
            new_key = re.sub(invalid_python_identifiers, replacement_str, new_key)
            if re.search(digit_start, new_key):
                new_key = ''.join([replacement_str, new_key])
            if not re.search(valid_python_identifer, new_key):
                raise FieldsError('Unable to scrub identity field into a valid python identifier [{}]'.format(k))
        identity_attributes[k] = new_key


class RootDataFields(IdentityFields):
    def __init__(self, additional_fields=None):
        self.base_fields = ['Session#',
                            'DeathSession',
                            'TipLivStatus',
                            'NumberOfTips',
                            'Date',
                            'Order']

        self.required_attributes = {k: v for k, v in IdentityFields.identity_attributes.items()}

        for k in self.base_fields:
            new_key = str(k)
            if not re.search(valid_python_identifer, new_key):
                new_key = re.sub(invalid_python_identifiers, replacement_str, new_key)
                if re.search(digit_start, new_key):
                    new_key = ''.join([replacement_str, new_key])
                if not re.search(valid_python_identifer, new_key):
                    raise FieldsError('Unable to scrub required field into a valid python identifier [{}]'.format(k))
            self.required_attributes[k] = new_key

        self.additional_fields = {}
        self.custom_attributes = {}

        if additional_fields:
            self.additional_fields = additional_fields
            for k, v in self.additional_fields.items():
                log.info('Preparing to extract custom field [{}]'.format(k))
                if k in self.required_attributes:
                    raise FieldsError('Additional field duplicates a required field [{}]'.format(k))
                if v not in [ROOT_BIRTH, ROOT_FINAL]:
                    raise FieldsError('Unknown custom field propogation value [{}][{}]'.format(k, v))
                new_key = str(k)
                if not re.search(valid_python_identifer, new_key):
                    new_key = re.sub(invalid_python_identifiers, replacement_str, new_key)
                    if re.search(digit_start, new_key):
                        new_key = ''.join([replacement_str, new_key])
                    if not re.search(valid_python_identifer, new_key):
                        raise FieldsError('Unable to scrub custom field into a valid python identifier [{}]'.format(k))
                self.custom_attributes[k] = new_key
                self.required_attributes[k] = new_key


class SynthesisDataFields(IdentityFields):
    def __init__(self):
        self.synthesis_fields = ['AliveTipsAtBirth',
                                 'AliveTipsAtDeath', ]
        self.required_attributes = {k: v for k, v in IdentityFields.identity_attributes.items()}

        for k in self.synthesis_fields:
            new_key = str(k)
            if not re.search(valid_python_identifer, new_key):
                new_key = re.sub(invalid_python_identifiers, replacement_str, new_key)
                if re.search(digit_start, new_key):
                    new_key = ''.join([replacement_str, new_key])
                if not re.search(valid_python_identifer, new_key):
                    raise FieldsError('Unable to scrub synthesis field into a valid python identifier [{}]'.format(k))
            self.required_attributes[k] = new_key

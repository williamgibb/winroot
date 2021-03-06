"""
Root container class.
"""
from __future__ import print_function
import collections
import logging

log = logging.getLogger(__name__)
__author__ = 'wgibb'


# Named tuples allow for the analyzer to create values for uniquely identifying roots if needed.
RootIdentity = collections.namedtuple('RootIdentity', ['rootname', 'location', 'birthsession'])


class Root(object):
    fixed_attributes = ['isAlive', 'censored', 'highestOrder', 'anomaly']

    def __init__(self, attr_map, rootname, location, birthsession):
        self.attr_map = attr_map
        self.identity = RootIdentity(rootname=rootname, location=location, birthsession=birthsession)
        self.anomaly = ''
        self.isAlive = ''
        self.censored = ''
        self.highestOrder = ''

    def set(self, key, value):
        new_key = self.attr_map.get(key, None)
        setattr(self, new_key, value)

    def get(self, key, default=None):
        new_key = self.attr_map.get(key)
        return getattr(self, new_key, default)

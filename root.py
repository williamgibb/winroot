# XXX Fill out docstring!
"""

"""
from __future__ import print_function
import collections
import logging

log = logging.getLogger(__name__)
__author__ = 'wgibb'


# Named tuples allow for the analyzer to create values for uniquely identifying roots if needed.
RootIdentity = collections.namedtuple('RootIdentity', ['rootname', 'location', 'birthsession'])
TipIdentity = collections.namedtuple('TipIdentity', ['birthsession', 'location'])


class Root(object):
    PROTECTED_FIELDS = ['rootName',
                        'location',
                        'birthsesssion',
                        'identity',
                        'tipIdentity',
                        'tube',
                        'session',
                        'goneSession',
                        'anomaly',
                        'order',
                        'aliveTipsAtBirth',
                        'aliveTipsAtGone',
                        'isAlive',
                        'PROTECTED_FIELDS']
    def __init__(self, rootname, location, birthsession):
        self.rootName = rootname
        self.location = location
        self.birthSession = birthsession
        self.identity = RootIdentity(rootname=rootname, location=location, birthsession=birthsession)
        self.tipIdentity = TipIdentity(birthsession=birthsession, location=location)
        self.date = ''
        self.tube = ''
        self.session = ''
        self.goneSession = ''
        self.anomaly = ''
        self.order = ''
        self.censored = ''
        self.aliveTipsAtBirth = ''
        self.aliveTipsAtGone = ''
        self.isAlive = ''

    def set_custom_field(self, key, value, force=False):
        if key in self.PROTECTED_FIELDS and not force:
            raise AttributeError('Cannot set a protected field as a custom value')
        setattr(self, key, value)
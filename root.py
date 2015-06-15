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
    def __init__(self, rootname, location, birthsession):
        self.rootName = rootname
        self.location = location
        self.birthSession = birthsession
        self.identity = RootIdentity(rootname=rootname, location=location, birthsession=birthsession)
        self.tipIdentity = TipIdentity(birthsession=birthsession, location=location)
        self.tube = ''
        self.session = ''
        self.goneSession = ''
        self.anomaly = ''
        self.order = ''
        self.highestOrder = ''
        self.censored = ''
        self.aliveTipsAtBirth = ''
        self.aliveTipsAtGone = ''
        self.avgDiameter = ''
        self.isAlive = ''
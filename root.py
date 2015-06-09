# XXX Fill out docstring!
"""

"""
from __future__ import print_function
import logging

log = logging.getLogger(__name__)
__author__ = 'wgibb'


class Root(object):
    def __init__(self, rootname, location, birthsession):
        self.rootName = rootname
        self.location = location
        self.birthSession = birthsession
        self.identity = (rootname, location, birthsession)
        self.tipIdentity = (birthsession, location)
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
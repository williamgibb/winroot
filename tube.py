# XXX Fill out docstring!
"""

"""
from __future__ import print_function
import logging

log = logging.getLogger(__name__)
__author__ = 'wgibb'


class Tube(object):
    def __init__(self, tubenumber):
        self.tubeNumber = tubenumber
        self.roots = []
        self.maxSessionCount = 0
        self.tipStats = ''
        self.sessionDates = {}
        self.index = 0

    def __iter__(self):
        for root in self.roots:
            yield root

    def __len__(self):
        return len(self.roots)

    def add_root(self, root):
        root.tube = self.tubeNumber
        self.roots.append(root)

    def insert_or_update_root(self, root):
        # insert root if the root identity is new
        insert = True
        for existingRoot in self:
            if root.identity == existingRoot.identity:
                insert = False
                # If there was a change, update the root attributes
                if existingRoot.isAlive.startswith('A') and root.isAlive.startswith('G'):
                    # root changed from A to G
                    log.debug('Changing root from A to G')
                    existingRoot.goneSession = root.goneSession
                    existingRoot.isAlive = root.isAlive
                elif existingRoot.isAlive.startswith('G') and root.isAlive.startswith('A'):
                    # root changed from G to A
                    log.debug('Changing root from G to A')
                    existingRoot.goneSession = ''
                    existingRoot.isAlive = root.isAlive
        # add the root to the tube
        if insert:
            # possible to insert a root at the last session.  likely rare though.
            # need to finalize this root before adding it into the tube.
            log.debug('Adding root to tube %s' % (str(root.identity)))
            self.add_root(root)
        return True

    def finalize_root(self, root):
        finalized = False
        for existingRoot in self:
            if root.identity == existingRoot.identity:
                if root.session == self.maxSessionCount:
                    existingRoot.highestOrder = root.order
                if existingRoot.isAlive.startswith('A'):
                    existingRoot.goneSession = 0
                    existingRoot.censored = 1
                if existingRoot.isAlive.startswith('G'):
                    existingRoot.censored = 0
                finalized = True
        return finalized
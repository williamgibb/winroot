# XXX Fill out docstring!
"""

"""
from __future__ import print_function
import logging
import fields

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
        root.set('Tube#', self.tubeNumber)
        self.roots.append(root)

    def insert_or_update_root(self, root):
        # insert root if the root identity is new
        insert = True
        for existingRoot in self:
            if root.identity == existingRoot.identity:
                insert = False
                # If there was a change, update the root attributes
                if existingRoot.isAlive.startswith('A') and root.isAlive.startswith(('G', 'D')):
                    # root changed from A to G
                    log.debug('Changing root from A to {}'.format(root.isAlive))
                    existingRoot.set('DeathSession', root.get('DeathSession'))
                    existingRoot.isAlive = root.isAlive
                elif existingRoot.isAlive.startswith(('G', 'D')) and root.isAlive.startswith('A'):
                    # root changed from G to A
                    log.debug('Changing root from {} to {}'.format(existingRoot.isAlive, root.isAlive))
                    existingRoot.set('DeathSession', '')
                    existingRoot.isAlive = root.isAlive
        # add the root to the tube
        if insert:
            # possible to insert a root at the last session.  likely rare though.
            # need to finalize this root before adding it into the tube.
            log.debug('Adding root to tube %s' % (str(root.identity)))
            self.add_root(root)
        return True

    def finalize_root(self, root_obj, root_fields):
        finalized = False
        for existingRoot in self:
            if root_obj.identity == existingRoot.identity:
                if root_obj.get('Session#') == self.maxSessionCount:
                    existingRoot.highestOrder = root_obj.get('Order')
                if existingRoot.isAlive.startswith('A'):
                    existingRoot.set('DeathSession', 0)
                    existingRoot.censored = 1
                # XXX Icky!
                if existingRoot.isAlive.startswith(('D', 'G')):
                    existingRoot.censored = 0
                # Update custom fields which are set when the root is finalized
                for attr, state in root_fields.additional_fields.items():
                    if state != fields.ROOT_FINAL:
                        continue
                    existingRoot.set(attr, root_obj.get(attr))
                finalized = True

        return finalized

    def insert_synthesis_data(self, sdata):
        for root_obj in self:
            # log.debug('Inserting synthesis data for {}'.format(root_obj.identity))
            sd = sdata.get(root_obj.identity)
            # log.debug(sd)
            for k, v in sd.items():
                root_obj.set(k, v)

from __future__ import absolute_import

import unittest
import logging

import win32com.client
from .xtn import IDualApplicationExtensions, IGroupExtensions

class TestComExtensionIntegration(unittest.TestCase):

    def Setup(self):
        logging.basicConfig(level=logging.DEBUG)

    def test_sanity(self):
        self.assertTrue(True)

    def test_esprit_scenario(self):
        IDualApplicationExtensions.register()
        esprit = win32com.client.Dispatch("Esprit.Application")
        self.assertTrue(hasattr(esprit, 'test_attribute'))
        self.assertFalse(hasattr(esprit, 'false_attribute'))

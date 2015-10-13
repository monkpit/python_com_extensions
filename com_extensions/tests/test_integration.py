from __future__ import absolute_import

try:
    import unittest2 as unittest
except ImportError:
    import unittest

import logging
import sys

import win32com.client
import win32com.client.makepy
from xtn import IShellDispatch5Extensions, MultiSubmoduleExtensions

class TestComExtensionIntegration(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        log_format = '%(asctime)-15s %(clientip)s %(user)-8s %(message)s'
        logging.basicConfig(format=log_format, level=logging.DEBUG)

    def test_sanity(self):
        self.assertTrue(IShellDispatch5Extensions is not None)

    def test_shell_scenario_single_submodule_extension(self):
        IShellDispatch5Extensions.register()
        shell_obj = win32com.client.Dispatch("Shell.Application")
        self.assertTrue(hasattr(shell_obj, 'test_attribute'))
        self.assertFalse(hasattr(shell_obj, 'false_attribute'))

    def test_shell_scenario_multiple_submodule_extensions(self):
        MultiSubmoduleExtensions.register()
        shell_obj = win32com.client.Dispatch("Shell.Application")
        self.assertTrue(hasattr(shell_obj, 'something1'))
        self.assertTrue(hasattr(shell_obj, 'something2'))
        shell_module = win32com.client.gencache.GetModuleForCLSID("{62112AA2-EBE4-11CF-A5FB-0020AFE7292D}")
        dsfve = shell_module.DShellFolderViewEvents()
        self.assertTrue(dsfve is not None)
        self.assertTrue(hasattr(dsfve, 'something1'))
        self.assertTrue(hasattr(dsfve, 'something2'))

from __future__ import absolute_import

try:
    import unittest2 as unittest
except ImportError:
    import unittest

import logging
import sys

import win32com.client
import win32com.client.makepy

from com_extensions import ComExtension, ComExtensionError

class TestComExtension(unittest.TestCase):
    @classmethod
    def setUp(cls):
        pass

    def test_sanity(self):
        self.assertTrue(ComExtension is not None)
        self.assertTrue(ComExtensionError is not None)

    def test_raises_ComExtensionError_with_no_progid(self):
        with self.assertRaisesRegexp(ComExtensionError, "When subclassing ComExtension"):
            ComExtension_with_no_progid.register()

    def test_raises_ComExtensionError_with_bad_progid(self):
        with self.assertRaisesRegexp(ComExtensionError, "Does it need to be generated"):
            ComExtension_with_bad_progid.register()

    def test_raises_ComExtensionError_with_no_extends(self):
        with self.assertRaisesRegexp(ComExtensionError, "'extends' must contain a list of COM classes"):
            ComExtension_with_no_extends.register()

    def test_raises_ComExtensionError_with_bad_single_submodule(self):
        with self.assertRaisesRegexp(ComExtensionError, "did not contain submodule"):
            ComExtension_with_bad_single_submodule.register()

    def test_raises_ComExtensionError_with_bad_extends(self):
        with self.assertRaisesRegexp(ComExtensionError, "did not contain submodule"):
            ComExtension_with_bad_extends.register()


class ComExtension_with_no_progid(ComExtension):
    progid = ''
    extends = ['submodule']


class ComExtension_with_bad_progid(ComExtension):
    progid = 'noodles'
    extends = ['submodule']


class ComExtension_with_no_extends(ComExtension):
    progid = 'Shell.Application'
    extends = None


class ComExtension_with_bad_single_submodule(ComExtension):
    progid = 'Shell.Application'
    extends = ['INoodles']


class ComExtension_with_bad_extends(ComExtension):
    progid = 'Shell.Application'
    extends = ['INoodles', 'INoodles2']

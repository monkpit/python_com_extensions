from __future__ import absolute_import, print_function

import win32com.client
import win32com.client.makepy
import sys
import inspect
import logging

logger = logging.getLogger(__name__)

class ComExtensionError(AttributeError):
    pass

class ComExtension(object):
    progid = None
    extends = ['']

    @classmethod
    def register(cls):
        if cls.progid is None or cls.progid == '':
            raise ComExtensionError("When subclassing ComExtension, {classname}.progid cannot be {value}".format(classname=cls.__name__, value=cls.progid))

        methods = [(name, method) for name, method \
                    in inspect.getmembers(cls, predicate=inspect.isroutine) \
                    if (name, method) not in inspect.getmembers(ComExtension) \
                    and name != '__subclasshook__' \
                    and name != 'register']

        logger.debug("Found methods: {methods}".format(methods=methods))

        if cls.extends is None \
                or cls.extends == [] \
                or cls.extends == [None]:
            cls.extends = ['']

        if cls.extends == ['']:
            raise ComExtensionError("'extends' must contain a list of COM classes defined in <<clsid>>.py")
        else:

            module = win32com.client.gencache.GetModuleForProgID(cls.progid)

            if module is None:
                logger.debug("Module for progid = {progid} not found.\nAttempting to generate it with makepy.py...".format(progid=cls.progid))
                oldargs = sys.argv
                sys.argv = ['', cls.progid]
                win32com.client.makepy.main()
                sys.argv = oldargs

                module = win32com.client.gencache.GetModuleForProgID(cls.progid)

                if module is None:
                    raise ComExtensionError("ComExtension error: module for progid '{progid}' was not found. " \
                                            "(Does it need to be generated with win32com/client/makepy.py ?)" \
                                            .format(progid=cls.progid))
                else:
                    logger.debug("Successfully created file for {progid}.".format(progid=cls.progid))

        for submodule_name in cls.extends:
            submodule = getattr(module, submodule_name, None)
            if submodule is None:
                raise ComExtensionError("ComExtension error: module '{module}' did not contain submodule '{submodule_name}'.".format(module=module, submodule_name=submodule_name))
            for name, method in methods:
                setattr(submodule, name, method)
                logger.debug("Added method {method}\n\tas '{name}'\n\tto module {module}.".format(method=method, name=name, module=submodule))

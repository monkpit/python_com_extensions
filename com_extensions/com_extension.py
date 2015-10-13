from __future__ import absolute_import, print_function

import win32com.client
import inspect
import logging

class ComExtension(object):
    progid = None
    submodules = ['']

    @classmethod
    def register(cls):
        if cls.progid is None:
            raise AttributeError("ComExtension error: progid cannot be 'None'")

        methods = [(name, method) for name, method in inspect.getmembers(cls, predicate=inspect.isfunction)]
        logging.debug("Found methods: {methods}".format(methods=methods))

        if cls.submodules is None or cls.submodules == [] or cls.submodules == [None]:
            cls.submodules = ['']

        if cls.submodules == ['']:
            raise AttributeError("ComExtension error: submodules must contain a list of COM classes defined in <<clsid>>.py")
        else:
            module = win32com.client.gencache.GetModuleForProgID(cls.progid)


        for submodule_name in cls.submodules:
            submodule = getattr(module, submodule_name, None)
            if submodule is None:
                raise AttributeError("ComExtension error: module '{module}' did not contain submodule '{submodule}'.".format(module=module, submodule=submodule))
            for name, method in methods:
                setattr(submodule, name, method)
                logging.debug("Added method {method}\n\tas '{name}'\n\tto module {module}.".format(method=method, name=name, module=submodule))

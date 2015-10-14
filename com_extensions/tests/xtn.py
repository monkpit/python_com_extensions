from com_extensions import ComExtension

class IGroupExtensions(ComExtension):
    progid = "Esprit.Application"
    extends = ["IGroup"]

    @staticmethod
    def __iter__(self):
        current = 1
        while current <= self.Count:
            yield self.Item(current)
            current += 1

    @staticmethod
    def __getitem__(self, k):
        return self.Item(k)

class IDualApplicationExtensions(ComExtension):
    progid = "Esprit.Application"
    extends = ["IDualApplication"]

    @staticmethod
    def test_attribute(self):
        return self

class IShellDispatch5Extensions(ComExtension):
    progid = "Shell.Application"
    extends = ["IShellDispatch5"]

    @staticmethod
    def test_attribute(self):
        return self

class MultiSubmoduleExtensions(ComExtension):
    progid = "Shell.Application"
    extends = ["IShellDispatch5", "DShellFolderViewEvents"]

    @staticmethod
    def something1(self):
        return self

    @staticmethod
    def something2(self):
        return self

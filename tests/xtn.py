from com_extensions import ComExtension

class IGroupExtensions(ComExtension):
    progid = "Esprit.Application"
    submodules = ["IGroup"]

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
    submodules = ["IDualApplication"]

    @staticmethod
    def test_attribute(self):
        pass

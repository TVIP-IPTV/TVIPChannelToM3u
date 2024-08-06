

class Config:
    __version__ = '0.0.1a1'

    def __init__(self) -> None:
        pass

    def is_alpha_version(self):
        return 'a' in self.__version__ or 'alpha' in self.__version__.lower()
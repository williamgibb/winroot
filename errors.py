"""
Exceptions
"""
__author__ = 'wgibb'


class AnalyzerError(Exception):
    pass


class DataError(AnalyzerError):
    pass


class SerializationError(AnalyzerError):
    pass


class FieldsError(Exception):
    pass

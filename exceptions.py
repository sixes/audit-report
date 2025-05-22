# Custom exception for invalid TB sheet format
class InvalidTBSheetFormatError(Exception):
    pass

# Custom exception for unrecognized items in TB sheet
class UnrecognizedItemError(Exception):
    pass

# Custom exception for invalid item names
class InvalidItemNameError(Exception):
    pass

# Custom exception for net assets and total equity mismatch
class NetAssetsEquityMismatchError(Exception):
    pass
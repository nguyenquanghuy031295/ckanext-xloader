from collections import defaultdict
from messytables.types import StringType, IntegerType, DateUtilType, BoolType
from openpyxl.cell.cell import TYPE_ERROR, TYPE_STRING, TYPE_NUMERIC, TYPE_BOOL

def column_count_modal(rows):
    """ Return the modal value of columns in the row_set's
    sample. This can be assumed to be the number of columns
    of the table. """
    counts = defaultdict(int)
    for row in rows:
        length = len([c for c in row if c.value])
        if length > 1:
            counts[length] += 1
    if not len(counts):
        return 0
    return max(list(counts.items()), key=lambda k_v: k_v[1])[0]

def headers_guess(rows, tolerance=1):
    """ Guess the offset and names of the headers of the row set.
    This will attempt to locate the first row within ``tolerance``
    of the mode of the number of columns in the row set sample.
    The return value is a tuple of the offset of the header row
    and the names of the columns.
    """
    rows = list(rows)
    modal = column_count_modal(rows)
    for i, row in enumerate(rows):
        length = len([c for c in row if c.value])
        if length >= modal - tolerance:
            # TODO: use type guessing to check that this row has
            # strings and does not conform to the type schema of
            # the table.
            return i, [u'{}'.format(c.value) for c in row if c.value]
    return 0, []

OPENPYXL_TYPE_MAPPING = {
  TYPE_STRING: StringType(),
  TYPE_BOOL: BoolType(),
  TYPE_NUMERIC: IntegerType(),
  'd': DateUtilType()
}

def _get_type_weight(type_str):
  if type_str == TYPE_STRING:
    return StringType.guessing_weight

  if type_str == TYPE_NUMERIC:
    return IntegerType.guessing_weight # Use "IntegerType" weight for even decimal + float

  if type_str == TYPE_BOOL:
    return BoolType.guessing_weight

  if type_str == 'd':
    return DateUtilType.guessing_weight

  return 0

def _map_openpyxl_type_to_messytable_type(type_str):
  type = OPENPYXL_TYPE_MAPPING.get(type_str, None)
  if type:
    return type
  
  return StringType()

def type_guess(cols, strict=False, header_offset=None):
  guesses = []
  if strict:
    for _, col in enumerate(cols):
      typesdict = {}
      for ci, cell in enumerate(col):
        if header_offset is not None and ci == header_offset:
          continue

        if not cell.value or cell.data_type == TYPE_ERROR:
          continue
        typesdict[cell.data_type] = 0
      
      has_item = bool(typesdict)
      if not has_item:
        typesdict[TYPE_STRING] = 0
      guesses.append(typesdict)
    
    for i, guess in enumerate(guesses):
      for type in guess:
        guesses[i][type] = _get_type_weight(type)
  else:
    for _, col in enumerate(cols):
      typesdict = {}
      for ci, cell in enumerate(col):
        if header_offset is not None and ci == header_offset:
          continue
        if not cell.value or cell.data_type == TYPE_ERROR:
          continue
        weight = typesdict.get(cell.data_type, 0) + _get_type_weight(cell.data_type)
        typesdict[cell.data_type] = weight
      
      has_item = bool(typesdict)
      if not has_item:
        typesdict[TYPE_STRING] = 0
      guesses.append(typesdict)

  _columns = []
  for guess in guesses:
    type_max_weight = max(guess, key=guess.get)
    _columns.append(_map_openpyxl_type_to_messytable_type(type_max_weight))
  
  return _columns
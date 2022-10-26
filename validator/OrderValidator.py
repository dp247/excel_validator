import functools
from validator.BaseValidator import BaseValidator


class OrderValidator(BaseValidator):
    message = "The order is not valid"

    def validate(self, value):
        # value = items from Excel doc
        # ext_items = items from YAML file
        value = super(OrderValidator, self).validate(value)
        if functools.reduce(lambda x, y: x and y, map(lambda p, q: p == q, value, self.ext_items), True):
            return True
        return False

    def __init__(self, params):
        super(OrderValidator, self).__init__(params)

        if 'items' not in params:
            raise ValueError("Header values not found in worksheet.")

        self.ext_items = params.get('items')

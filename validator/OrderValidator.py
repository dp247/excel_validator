from validator.BaseValidator import BaseValidator

class OrderValidator(BaseValidator):
    message = "The order is not valid"

    def validate(self, values):
        # possible null values
        # value = super(OrderValidator, self).validate(value)
        for value in values:


    def __init__(self, params):
        super(OrderValidator, self).__init__(params)

        if not 'items' in params:
            raise ValueError("Item order not set")
        self.choices = params.get('items')

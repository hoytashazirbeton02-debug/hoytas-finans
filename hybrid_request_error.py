class RequestError(Exception):
    """A request body failed HTTP input validation."""

    def __init__(self, status):
        super().__init__(status)
        self.status = status

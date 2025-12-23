class IscConnectorError(Exception):
    pass


class SeminarDownloaderError(IscConnectorError):
    pass


class SeminarDownloaderHttpError(IscConnectorError):
    status_code: int

    def __init__(self, status_code, *args, **kwargs):
        self.status_code = status_code
        super().__init__(args, kwargs)


class SeminarDownloaderNotFound(SeminarDownloaderHttpError):
    def __init__(self, *args, **kwargs):
        self.status_code = 404
        super().__init__(args, kwargs)

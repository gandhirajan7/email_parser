class Feedback:

    __slots__ = 'fNo', 'name', 'mail', 'product', 'review'

    def __init__(self, fid, name, mail, product, review):
        self.fNo = fid
        self.name = name
        self.mail = mail
        self.product = product
        self.review = review

    def __str__(self):
        return f"{self.fNo}, {self.name}, {self.mail}, {self.product}, {self.review}"

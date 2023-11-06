from pymongo import MongoClient


class Database:
    def __init__(self):
        self.cluster = MongoClient('mongodb://localhost:27017')
        self.db = self.cluster.dbTesi

    def insert_data(self, data):
        collection = self.db.pdfData
        result = collection.insert_one(data)
        if result:
            return True
        else:
            return False

    def insert_rule(self, rule):
        collection = self.db.Regole
        result = collection.insert_one(rule)
        if result:
            return True
        else:
            return False

    def get_rules(self):
        collection = self.db.Regole
        result = collection.find({}, {"_id": 0})
        return result

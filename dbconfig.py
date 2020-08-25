from pymongo import MongoClient

client = MongoClient('mongodb+srv://ss_db:passss@cluster0.8chqj.mongodb.net/test?retryWrites=true&w=majority')

def dbConnect():
    return client
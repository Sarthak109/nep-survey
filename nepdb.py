import dbconfig
import json

client= dbconfig.dbConnect()

db=client['NEP']

def  add_survey(survey):
    col=db['Survey']
    col.insert_one(survey)

def update_questions(topic_id,expression):
    col=db['questions']
    col.update({'topicId':topic_id},{"$set":{"data":expression}})


def add_question(topic_id,expression):
    col=db['questions']
    col.update({'topicId':topic_id},{"$push":{"data":expression}})


def delete_question(topic_id,ref):
    col=db['questions']
    col.update({'topicId':topic_id},{"$pull":{"data":{"ref":ref}}})

def add_section(doc):
    col=db['questions']
    col.insert(doc)

def delete_section(topic_id):
    col=db['questions']
    col.delete_one({'topicId':topic_id})

def get_questions():
    col=db['questions']
    q = []
    for x in col.find():
        x['_id']=str(x['_id'])
        q.append(x)
    return json.dumps(q)

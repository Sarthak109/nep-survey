
from flask import Flask,jsonify, request, send_file
from flask_restful import Api,Resource
from flask.json import JSONEncoder
import nepdb as dbs
from security import authenticate,identity
from flask_jwt import JWT, jwt_required,current_identity
from flask_cors import CORS
import json
import dbconfig
import xlsxwriter

client= dbconfig.dbConnect()

db=client['NEP']


app=Flask(__name__)
app.config['SECRET_KEY'] = 'app@123!'
jwt = JWT(app,authenticate,identity)
api = Api(app)
CORS(app)



@app.route("/addsurvey",methods=["POST"])
def add_new_survey():
    surveys=request.json
    dbs.add_survey(surveys)
    return {"message":"done"}

@app.route("/getquestions")
def getquestions():
    return dbs.get_questions()

@app.route("/updatequestion/<string:topic_id>",methods=["PUT"])
def updatequestion(topic_id):
    expression=request.json
    dbs.update_questions(topic_id,expression)
    return {"message":"done"}


@app.route("/addquestion/<string:topic_id>",methods=["PUT"])
def addquestion(topic_id):
    expression=request.json
    dbs.add_question(topic_id,expression)
    return {"message":"done"}

@app.route("/deletequestion/<string:topic_id>/<string:ref>",methods=["DELETE"])
def deletequestion(topic_id,ref):
    dbs.delete_question(topic_id,ref)
    return {"message":"done"}

@app.route("/addsection",methods=["POST"])
def addsection():
    tid=request.json["topicId"]
    tname=request.json["topicName"]
    doc={"topicId":tid,"topicName":tname,"data":[]}
    dbs.add_section(doc)
    return {"message":"done"}

@app.route("/deletesection/<string:topic_id>",methods=["DELETE"])
def deletesection(topic_id):
    dbs.delete_section(topic_id)
    return {"message":"done"}


@app.route('/getAllIds',methods = ["GET"])
def getAllIds():
    ret = []
    result = db['questions'].aggregate([
    {
        '$project': {
            'topicId': 1, 
            '_id': 0
        }
    }
    ])
    for re in result:
        ret.append(re['topicId'])
    return {"message":ret}

@app.route('/getAllhds',methods = ["GET"])
def getAllhds():
    ret = []
    result = db['questions'].aggregate([
    {
        '$project': {
            'topicName': 1, 
            '_id': 0
        }
    }
    ])
    for re in result:
        ret.append(re['topicName'])
    return {"message":ret}


@app.route('/getsubs',methods = ["POST"])
def getsubs():
    para = request.json
    ret = []
    result = db['questions'].aggregate([
        {
            '$match': {
                'topicId': para["id"]
            }
        }, {
            '$unwind': {
                'path': '$data'
            }
        }, {
            '$project': {
                'data.desc': 1, 
                '_id': 0
            }
        }
    ])
    for re in result:
        ret.append(re["data"]["desc"])
    return {"message":ret}

@app.route('/getsubsbyhd',methods = ["POST"])
def getsubsbyhd():
    para = request.json
    ret = []
    result = db['questions'].aggregate([
        {
            '$match': {
                'topicName': para["hd"]
            }
        }, {
            '$unwind': {
                'path': '$data'
            }
        }, {
            '$project': {
                'data.desc': 1, 
                '_id': 0
            }
        }
    ])
    for re in result:
        ret.append(re["data"]["desc"])
    return {"message":ret}

@app.route('/getchartsid',methods = ["POST"])
def getchartsid():
    para = request.json
    idname = para["id"]
    sec1 = db['Survey'].aggregate([
        {
            '$unwind': {
                'path': '$answers'
            }
        }, {
            '$match': {
                'answers.topicId': idname,
                'answers.section': 1
            }
        }, {
            '$group': {
                '_id': '$answers.choice', 
                'total': {
                    '$sum': 1
                }
            }
        }
    ])
    sec2 = db['Survey'].aggregate([
        {
            '$unwind': {
                'path': '$answers'
            }
        }, {
            '$match': {
                'answers.topicId': idname,
                'answers.section': 2
            }
        }, {
            '$group': {
                '_id': '$answers.choice', 
                'total': {
                    '$sum': 1
                }
            }
        }
    ])
    sec3 = db['Survey'].aggregate([
        {
            '$unwind': {
                'path': '$answers'
            }
        }, {
            '$match': {
                'answers.topicId': idname,
                'answers.section': 3
            }
        }, {
            '$group': {
                '_id': '$answers.choice', 
                'total': {
                    '$sum': 1
                }
            }
        }
    ])
    sd1 = []
    for x in sec1:
        sd1.append({"option":x["_id"],"count":x["total"]})
    sd2 = []
    for x in sec2:
        sd2.append({"option":x["_id"],"count":x["total"]})
    sd3 = []
    for x in sec3:
        sd3.append({"option":x["_id"],"count":x["total"]})
    
    return {"sd1": sd1,"sd2":sd2,"sd3":sd3}




@app.route('/getchartshd',methods = ["POST"])
def getchartshd():
    para = request.json
    result = db['questions'].aggregate([
        {
            '$match': {
                'topicName': para['hd']
            }
        }
    ])
    idres = ""
    for re in result:
        print(re['topicId'])
        idres = re['topicId']
    sec1 = db['Survey'].aggregate([
        {
            '$unwind': {
                'path': '$answers'
            }
        }, {
            '$match': {
                'answers.topicId': idres,
                'answers.section': 1
            }
        }, {
            '$group': {
                '_id': '$answers.choice', 
                'total': {
                    '$sum': 1
                }
            }
        }
    ])
    sec2 = db['Survey'].aggregate([
        {
            '$unwind': {
                'path': '$answers'
            }
        }, {
            '$match': {
                'answers.topicId': idres,
                'answers.section': 2
            }
        }, {
            '$group': {
                '_id': '$answers.choice', 
                'total': {
                    '$sum': 1
                }
            }
        }
    ])
    sec3 = db['Survey'].aggregate([
        {
            '$unwind': {
                'path': '$answers'
            }
        }, {
            '$match': {
                'answers.topicId': idres,
                'answers.section': 3
            }
        }, {
            '$group': {
                '_id': '$answers.choice', 
                'total': {
                    '$sum': 1
                }
            }
        }
    ])
    sd1 = []
    for x in sec1:
        sd1.append({"option":x["_id"],"count":x["total"]})
    sd2 = []
    for x in sec2:
        sd2.append({"option":x["_id"],"count":x["total"]})
    sd3 = []
    for x in sec3:
        sd3.append({"option":x["_id"],"count":x["total"]})
    
    return {"sd1": sd1,"sd2":sd2,"sd3":sd3}
    



@app.route('/getsubchart',methods = ["POST"])
def subchart():
    para = request.json
    result = db['questions'].aggregate([
        {
            '$unwind': {
                'path': '$data'
            }
        }, {
            '$match': {
                'data.desc': para['sub']
            }
        }, {
            '$project': {
                'data.ref': 1, 
                '_id': 0
            }
        }
    ])
    ref = ""
    for re in result:
        print(re["data"]["ref"])
        ref = re["data"]["ref"]
    sec1 = db['Survey'].aggregate([
        {
            '$unwind': {
                'path': '$answers'
            }
        }, {
            '$match': {
                'answers.ref': ref,
                'answers.section': 1
            }
        }, {
            '$group': {
                '_id': '$answers.choice', 
                'total': {
                    '$sum': 1
                }
            }
        }
    ])
    sec2 = db['Survey'].aggregate([
        {
            '$unwind': {
                'path': '$answers'
            }
        }, {
            '$match': {
                'answers.ref': ref,
                'answers.section': 2
            }
        }, {
            '$group': {
                '_id': '$answers.choice', 
                'total': {
                    '$sum': 1
                }
            }
        }
    ])
    sec3 = db['Survey'].aggregate([
        {
            '$unwind': {
                'path': '$answers'
            }
        }, {
            '$match': {
                'answers.ref': ref,
                'answers.section': 3
            }
        }, {
            '$group': {
                '_id': '$answers.choice', 
                'total': {
                    '$sum': 1
                }
            }
        }
    ])
    sd1 = []
    for x in sec1:
        sd1.append({"option":x["_id"],"count":x["total"]})
    sd2 = []
    for x in sec2:
        sd2.append({"option":x["_id"],"count":x["total"]})
    sd3 = []
    for x in sec3:
        sd3.append({"option":x["_id"],"count":x["total"]})
    
    return {"sd1": sd1,"sd2":sd2,"sd3":sd3}







@app.route('/getresct',methods = ["GET"])
def resct():
    result = db['Survey'].aggregate([
        {
            '$project': {
                'email': 1, 
                '_id': 0
            }
        }
    ])
    cnt = 0
    for re in result:
        cnt = cnt+1
    return {"message":cnt}

@app.route('/getorct',methods = ["GET"])
def orct():
    result = db['Survey'].aggregate([
        {
            '$group': {
                '_id': {
                    'org': '$org'
                }, 
                'cnt': {
                    '$sum': 1
                }
            }
        }
    ])
    cnt = 0
    for re in result:
        cnt = cnt+1
    return {"message":cnt}

@app.route('/getevy',methods = ["GET"])
def getevy():
    sec1 = db['Survey'].aggregate([
        {
            '$unwind': {
                'path': '$answers'
            }
        }, {
            '$match': {
                'answers.section': 1
            }
        }, {
            '$group': {
                '_id': '$answers.choice', 
                'total': {
                    '$sum': 1
                }
            }
        }
    ])
    sec2 = db['Survey'].aggregate([
        {
            '$unwind': {
                'path': '$answers'
            }
        }, {
            '$match': {
                'answers.section': 2
            }
        }, {
            '$group': {
                '_id': '$answers.choice', 
                'total': {
                    '$sum': 1
                }
            }
        }
    ])
    sec3 = db['Survey'].aggregate([
        {
            '$unwind': {
                'path': '$answers'
            }
        }, {
            '$match': {
                'answers.section': 3
            }
        }, {
            '$group': {
                '_id': '$answers.choice', 
                'total': {
                    '$sum': 1
                }
            }
        }
    ])
    sd1 = []
    for x in sec1:
        sd1.append({"option":x["_id"],"count":x["total"]})
    sd2 = []
    for x in sec2:
        sd2.append({"option":x["_id"],"count":x["total"]})
    sd3 = []
    for x in sec3:
        sd3.append({"option":x["_id"],"count":x["total"]})
    
    return {"sd1": sd1,"sd2":sd2,"sd3":sd3}

def subchsheet(desc):
    result = db['questions'].aggregate([
        {
            '$unwind': {
                'path': '$data'
            }
        }, {
            '$match': {
                'data.desc': desc
            }
        }, {
            '$project': {
                'data.ref': 1, 
                '_id': 0
            }
        }
    ])
    ref = ""
    for re in result:
        print(re["data"]["ref"])
        ref = re["data"]["ref"]
    sec1 = db['Survey'].aggregate([
        {
            '$unwind': {
                'path': '$answers'
            }
        }, {
            '$match': {
                'answers.ref': ref,
                'answers.section': 1
            }
        }, {
            '$group': {
                '_id': '$answers.choice', 
                'total': {
                    '$sum': 1
                }
            }
        }
    ])
    sec2 = db['Survey'].aggregate([
        {
            '$unwind': {
                'path': '$answers'
            }
        }, {
            '$match': {
                'answers.ref': ref,
                'answers.section': 2
            }
        }, {
            '$group': {
                '_id': '$answers.choice', 
                'total': {
                    '$sum': 1
                }
            }
        }
    ])
    sec3 = db['Survey'].aggregate([
        {
            '$unwind': {
                'path': '$answers'
            }
        }, {
            '$match': {
                'answers.ref': ref,
                'answers.section': 3
            }
        }, {
            '$group': {
                '_id': '$answers.choice', 
                'total': {
                    '$sum': 1
                }
            }
        }
    ])
    sd1 = []
    for x in sec1:
        sd1.append({"option":x["_id"],"count":x["total"]})
    sd2 = []
    for x in sec2:
        sd2.append({"option":x["_id"],"count":x["total"]})
    sd3 = []
    for x in sec3:
        sd3.append({"option":x["_id"],"count":x["total"]})
    
    return {"sd1": sd1,"sd2":sd2,"sd3":sd3}


@app.route('/report',methods = ["GET"])
def sheet():
    workbook = xlsxwriter.Workbook('reports.xlsx')
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})
    main = workbook.add_format()
    main.set_bold()
    main.set_bg_color('blue')
    main.set_font_color('white')
    main.set_underline()
    secf = workbook.add_format()
    secf.set_font_color('blue')
    secf.set_bold()
    secf.set_underline()
    num = workbook.add_format()
    num.set_align('center')
    worksheet.write(0,0,'REPORTS FOR THE SURVEY ON SCHOOL EDUCATION',main)
    row = 4
    col = 0
    ret = []
    result = db['questions'].aggregate([
    {
        '$project': {
            'topicName': 1, 
            '_id': 0
        }
    }
    ])
    for re in result:
        worksheet.write(row,col,re['topicName'],secf)
        row +=3
        col +=1
        worksheet.write(row-1,col+1,"Present Status in Karnataka",bold)
        worksheet.write(row-1,col+3,"Nature of Implications",bold)
        worksheet.write(row-1,col+5,"Implementation time",bold)
        res = db['questions'].aggregate([
            {
                '$match': {
                    'topicName': re['topicName']
                }
            }, {
                '$unwind': {
                    'path': '$data'
                }
            }, {
                '$project': {
                    'data.desc': 1, 
                    '_id': 0
                }
            }
        ])
        for desc in res:
            ret.append(desc["data"]["desc"])
            rss = row
            css = col+3
            rsss = row
            csss = col+5
            rs = row
            cs = col+1
            worksheet.write(row,col,desc["data"]["desc"])
            rsts =  subchsheet(desc["data"]["desc"])
            sd1 = rsts["sd1"]
            sd2 = rsts["sd2"]
            sd3 = rsts["sd3"]
            for op in sd1:
                rs +=1
                worksheet.write(rs,cs,op['option'])
                worksheet.write(rs,cs+1,op['count'],num)
            for op in sd2:
                rss+=1
                worksheet.write(rss,css,op['option'])
                worksheet.write(rss,css+1,op['count'],num)
            for op in sd3:
                rsss+=1
                worksheet.write(rsss,csss,op['option'])
                worksheet.write(rsss,csss+1,op['count'],num)

            row+=4
        row+=2
        col = 0
    workbook.close()
    return send_file('./reports.xlsx',attachment_filename="report.xlsx")


if __name__=='__main__':
    app.run(threaded=True, port=5000)


class User(object):

    def __init__(self, _id, username, password):
        self.id = _id
        self.username = username
        self.password = password

    def __str__(self):
        return f"User({self.username},'*******',{self.password})"


users_list = [
    User(1002,'krish','krish@123'),
]

user_by_username = { user.username:user  for user in users_list}
user_by_id       = { user.id:user  for user in users_list}

def authenticate(username,password):
    user = user_by_username.get(username,None)
    if user and user.password == password:
        return user 

def identity(playload):
    user_id = playload['identity']
    return user_by_id.get(user_id,None)
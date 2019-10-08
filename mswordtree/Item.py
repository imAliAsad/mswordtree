import json
from json import JSONEncoder
import uuid 


class Item(JSONEncoder):

    def __init__(self):
        self.GUID = str(uuid.uuid4())
        self.Type = ''        
        self.Content = ''
        self.N = None
        self.Parent = None
        self.Items = []
    
    def toJSON(self):
        return json.dumps(self, default=lambda obj: obj.__dict__)








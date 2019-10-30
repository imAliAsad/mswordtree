
def recurse(root, filename, nodes):    
    
    for child in root.GetContent():        
        obj = JsonObject_Item(filename, root, child)        
        nodes.append(obj)        

    for child in root.GetSubHeadings():
        obj = JsonObject_Item(filename, root, child)        
        nodes.append(obj)
        recurse(child, filename, nodes)

def ToString(roots):
    nodes = []
    for root in roots: 
        obj = JsonObject_Item(root.Content, root, root)
        obj["Parent"] = ''
        nodes.append(obj)
        recurse(root, root.Content, nodes)
    return nodes

def JsonObject_Item(filename, parent, child):

    obj = {"document": filename, "Parent": parent.GUID, "GUID": child.GUID, "Content": "", "TableContent": [], "ColumnNames": [],  "Type": child.Type, "Tags": [], "QA":[]}
        
    if (child.Type is "Table"):
        df = child.Content
        obj["ColumnNames"] = list(df.columns.values)
        Row_list =[]   
        # Iterate over each row 
        for i in range((df.shape[0])):
            # Using iloc to access the values of  
            # the current row denoted by "i" 
            Row_list.append(list(df.iloc[i, :]))  
        obj["TableContent"] = Row_list
        pass
    else:
        obj["Content"] = child.Content

    if hasattr(child, 'Tags'):
        obj["Tags"] = child.Tags
    if hasattr(child, 'QA'):
        obj["QA"] = child.QA

    return obj


def GetAllHeadings(item, filename):
    json = []
    for head in item.GetSubHeadings():
        obj = JsonObject_Item(filename,item,head)        
        items = GetAllHeadings(head, filename)
        obj["Items"] = items
        json.append(obj)
    return json
    

def ToString_AllHeadings(item):
    data = GetAllHeadings(item, item.Content)
    return data
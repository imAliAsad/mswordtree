    
def recurse(root, filename, nodes):
    children = root.Items
    if len(children) <= 0:
        return

    for child in children:        
        obj = JsonObject_Item(filename, root, child)        
        nodes.append(obj)
        recurse(child, filename, nodes)

    
def ToString(roots):
    nodes = []
    for root in roots:
        recurse(root, root.Content, nodes)
    return nodes

def JsonObject_Item(filename, parent, child):

    obj = {"document": filename, "Parent": parent.GUID, "GUID": child.GUID, "Content": "", "TableContent": [], "ColumnNames": [],  "Type": child.Type, "N": child.N}
        
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
    return obj
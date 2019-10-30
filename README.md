## mswordtree

Parse your whole word document in a hierarchical tree structure. The document content will be listed down as Heading and its children as subheading/paragraph/table etc.

Install the library using following comand

```
pip install mswordtree
```

Use the following code to parse your word document in a tree structure

```python
from mswordtree import GetWordDocTree
root = GetWordDocTree('test.docx')
```
Now you can iterate over all objects of the document by using the following code

```
for item in root.Items:
    print('Type: {} -> Content {}\n'.format(item.Type, item.Content))
```

To make the json use the following code

```python
from mswordtree import ToString
ToString([root])
```


### Common Methods

#### Find(guid)

Use the root element to find any element in its tree structure by mathing its GUID.

```python
item = root.Find('3b34509b-533e-40cc-b0dc-c44df5bcba51')
```

#### ToString_AllHeadings(root)

Returns the string of all heading elements in a tree structure, which we can use as a json string.

```python
from mswordtree import ToString_AllHeadings
import json

data = ToString_AllHeadings(root)
json.dumps(data)
```

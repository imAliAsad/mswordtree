## mswordtree

Parse your whole word document in a hierarchical tree structure. The document content will be listed down as Heading and its children as subheading/paragraph/table etc.

Install the library using following comand

```python
pip install mswordtree
```

Use the following code to parse your word document in a tree structure

```python
from mswordtree import GetWordDocTree
root = GetWordDocTree('test.docx')
```
Now you can iterate over all objects of the document by using the following code

```python
for item in root.Items:
    print('Type: {} -> Content {}\n'.format(item.Type, item.Content))
```

To make the json use the following code

```python
from mswordtree import ToString
ToString([root])
```

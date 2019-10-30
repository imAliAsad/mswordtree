from mswordtree import GetWordDocTree
from Helper import ToString



data = ToString([GetWordDocTree('LOA - Sub Consultant ArchCorp.docx')])
print(data)
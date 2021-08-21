# memtoxml
Simle utility to convert VFP memory from/to an XML document.

How to use:

#### Instantiate a MemToXml object:
```
  LOCAL loMemToXml
  loMemToXMl = MemToXml()
```

#### Restore variables from a file into memory:
```
  lbOk = loMemToXml.RestoreFromFile([XmlSaveFile],[IgnoreList])
```

#### Save memory to an XML file:
```
  lbOk = loMemToXml.SaveToFile([XmlSaveFile],[IgnoreList])
```

#### Restore variables to an object:
```
  lbOk = loMemToXml.RestoreObject([ObjectReference],[XmlSaveFile])
```

#### Save object properties to an XML file:
```
  lbOk = loMemToXml.SaveObject([ObjectReference],[XmlSaveFile])
```

##### Parameters
- XmlSaveFile - the file path and name of an XML memory file
- IgnoreList - comma-separated list of variable names to "ignore"
- ObjectReference - an instantiated object (e.g. <code>CREATOBJECT('EMPTY')</code>)

RETURN CREATEOBJECT('MemToXml')

DEFINE CLASS MemToXml AS Exception
	PROTECTED bObjectMode,oMemoryObject
	bObjectMode = .f.

	*  set to TRUE to output the variables sorted in ascending order by name
	bSortVariables = .t. && output from SaveObject() will always be sorted
	
	* a comma-separated list of variable prefixes (or names) to compare
	cExclusionList = ''

	bIgnoreLocal = .f.
	bIgnorePublic = .f.
	bIgnorePrivate = .f.
	bIgnoreHidden = .t.
	
	* restore variables into a memory object from the specified xml memory file
	* _tcIgnoreList is a comma separated list of prefixes to ignore
	PROCEDURE RestoreObject (_tcXmlFile, _tcIgnoreList)
		LOCAL lbOk
		this.bObjectMode = .t.
		this.oMemoryObject = CREATEOBJECT('EMPTY')

		lbOk = this.RestoreFromFile(m._tcXmlFile, m._tcIgnoreList)

		this.bObjectMode = .f.

		RETURN IIF(m.lbOk,this.oMemoryObject,.f.)
	ENDPROC

	* save the member properties of the passed in object to an xml memory file
	PROCEDURE SaveObject (toMemObject, tcXmlFile, tcIgnoreList)
		this._SetSettings()
		LOCAL lbOk,laMembers[1],lnMemberCount,laIgnore[1],lnIgnoreCount

		this.bObjectMode = .t.
		this.oMemoryObject = toMemObject

		IF EMPTY(m.tcIgnoreList)
			m.tcIgnoreList = this.cExclusionList
		ENDIF
		IF !EMPTY(m.tcIgnoreList)
			=ALINES(m.laIgnore,UPPER(m.tcIgnoreList),4,',')
		ENDIF

		* get a list of public user-defined properties
		lnMemberCount = AMEMBERS(laMembers,toMemObject,0,'G+U')

		LOCAL laEntries[lnMemberCount,5],lnStep,lnMemStep,lcVarName,lcValue,lcVariableType,loException

		lnMemStep = 1
		FOR lnStep = 1 TO lnMemberCount
			IF !EMPTY(m.tcIgnoreList) AND this._IgnoreVariable(@laIgnore,m.laMembers[m.lnStep])
				LOOP
			ENDIF
			lnMemStep = m.lnMemStep + 1
			laEntries[m.lnStep,1] = m.laMembers[m.lnStep]
			lcVarName = 'toMemObject.'+m.laMembers[m.lnStep]

			TRY
				IF TYPE(m.lcVarName,1)='A'
					* array property
					lcVariableType = 'A'
					StopHere(lcVarName)

					LOCAL lnRow,lnCol,lnArrayHeight,lnArrayWidth
					lnArrayHeight = ALEN(&lcVarName,1)
					lnArrayWidth = ALEN(&lcVarName,2)
					laEntries[m.lnStep,4] = m.lnArrayHeight
					laEntries[m.lnStep,5] = m.lnArrayWidth

					PRIVATE _paMemberArray
					IF lnArrayWidth > 0
						DIMENSION _paMemberArray[m.lnArrayHeight,m.lnArrayWidth]
					ELSE
						DIMENSION _paMemberArray[m.lnArrayHeight]
					ENDIF

					ACOPY(&lcVarName,m._paMemberArray)

					StopHere()
					
					lcValue = this._GetLiveValue('_paMemberArray','A','',m.lnArrayHeight,m.lnArrayWidth)
				ELSE
					lcVariableType = TYPE(m.lcVarName)
					lcValue = TRAN(EVAL(m.lcVarName))
				ENDIF
			CATCH TO loException
				LOCAL lcErrMess
				lcErrMess = 'MemToXml save variable error ('+m.lcVarName+'):'+CHR(13)+CHR(10)+m.loException.Message
				
				whereami(m.lcErrMess)
				WAIT WINDOW m.lcErrMess TIMEOUT .2
			ENDTRY

			laEntries[m.lnStep,2] = m.lcVariableType
			laEntries[m.lnStep,3] = m.lcValue
		ENDFOR

		this._WriteToXml(@laEntries, m.lnMemberCount, m.tcXmlFile)

		this.bObjectMode = .f.
		this._ClearSettings()
	ENDPROC

	* restore memory variables from an xml file
	* note that only private/public variables that are already defined in the calling program can be recovered
	PROCEDURE RestoreFromFile (_tcXmlFile,_tcIgnoreList)

		LOCAL _loXml,_lbOk,_loException
		_loXml = this._GetDomDoc()
		IF VARTYPE(_loXml)#'O'
			RETURN .f.
		ENDIF

		_loXml.setProperty("SelectionLanguage", "XPath")

		IF VARTYPE(m._tcXmlFile)#'C' OR EMPTY(m._tcXmlFile)
			_tcXmlFile = GETFILE('XML')
			IF EMPTY(m._tcXmlFile)
				RETURN .f.
			ENDIF
		ENDIF

		IF !FILE(m._tcXmlFile)
			RETURN .f.
		ENDIF

		TRY
			_loXml.LoadXml(FILETOSTR(m._tcXmlFile))
			_lbOk = .t.
		CATCH TO m._loException
			do erro with m._loException
		ENDTRY

		IF NOT m._lbOk
			RETURN .f.
		ENDIF

		LOCAL _loRoot,_lnStep,_lnCount,_loNode,_lcName,_lcType,_lcValue,_lcExecString,_laIgnore[1]
		_loRoot = m._loXml.selectSingleNode('//memory')
		IF VARTYPE(m._loRoot)#'O'
			LOCAL _lcMessage2
			IF m._loXml.parseError.Line > 0
				_lcMessage2 = 'Parse Error:'+m._loXml.parseError.Reason
			ELSE
				_lcMessage2 = 'File does not appear to be a valid xml memory file.'
			ENDIF
			do erro with 0,'XML memory file is improperly formatted.',m._lcMessage2,m._tcXmlFile,m._loXml.parseError.Line
			RETURN .f.
		ENDIF

		IF EMPTY(m._tcIgnoreList) AND !EMPTY(this.cExclusionList)
			_tcIgnoreLIst = this.cExclusionList
		ENDIF
		IF !EMPTY(m._tcIgnoreList)
			ALINES(m._laIgnore,UPPER(m._tcIgnoreList),4,',')
		ENDIF
		
		_lnCount = m._loRoot.childNodes.length
		FOR m._lnStep = 0 TO m._lnCount - 1
			_loNode = m._loRoot.childNodes.item(_lnStep)

			_lcName = m._loNode.getAttribute('name')
			_lcType = m._loNode.getAttribute('type')
			
			* handle a bug in the way variables are stored
			IF ISLOWER(_lcType)
				* some variables are listed in the LIST MEMORY output, when they are
				* hidden by other variables (by being passed as parameters)... in which case
				* the overriding variable name is in the place where the Type should be.
				LOOP
			ENDIF
						
			_lcValue = ''
			IF VARTYPE(m._loNode.firstChild)='O' AND m._loNode.firstChild.nodeName = '#'
				_lcValue = m._loNode.text
				IF m._loNode.firstChild.nodeName = '#cdata'
					* xml is standardized around using only the LF (chr(10)) instead of a CR+LF (CHR(13)+CHR(10))
					* which means that the xml reader returns LF instead of CRLF,
					* however Windows does not recognize a LF as a line break so we have to convert 
					IF NOT CHR(13)+CHR(10)$m._lcValue
						_lcValue = STRTRAN(m._lcValue,CHR(10),CHR(13)+CHR(10))
					ENDIF
				ENDIF
			ENDIF


			* TBD: fix this so that it works with the ignore list
			IF !EMPTY(m._tcIgnoreList) AND this._IgnoreVariable(@m._laIgnore,m._lcName)
				LOOP
			ENDIF

			_lcExecString = ''

			IF m._lcType = 'A'
				LOCAL _laElements[1,3],_lnArrayStep,_lnElementCount,_lnHeight,_lnWidth
				_lnHeight = VAL(NVL(m._loNode.getAttribute('height'),'1'))
				_lnWidth = VAL(NVL(m._loNode.getAttribute('width'),'0'))

				IF m._loNode.childNodes.length == 1 AND m._loNode.firstChild.nodeName='#'
					_lnElementCount = this._GetArrayFromString(@m._laElements,m._lcName,m._lcValue,@m._lnHeight,@m._lnWidth)
				ELSE
					_lnElementCount = this._GetArrayFromXml(@m._laElements,m._lcName,m._loNode,@m._lnHeight,@m._lnWidth)
				ENDIF
				*!* IF _lnElementCount > 0
					IF this.bObjectMode
						* add the array property to the object
						ADDPROPERTY(this.oMemoryObject,m._lcName+'[1]')
						m._lcName = 'this.oMemoryObject.'+m._lcName
					ENDIF	
					_lcExecString = 'DIMENSION '+m._lcName+'['+TRAN(m._lnHeight)+IIF(EMPTY(m._lnWidth),'',','+TRAN(m._lnWidth))+']'
					&_lcExecString

					STORE .f. TO (m._lcName)
					FOR m._lnArrayStep = 1 TO m._lnElementCount
						this._SetVariable(m._laElements[m._lnArrayStep,1],m._laElements[m._lnArrayStep,2],m._laElements[m._lnArrayStep,3])
					ENDFOR
				*!* ENDIF
			ELSE
				this._SetVariable(m._lcName,m._lcType,m._lcValue)
			ENDIF
		ENDFOR
		
	ENDPROC

	* save the current memory variables to an xml file
	* note that only private/public variables will be saved (because locals wont pass through to this function)
	PROCEDURE SaveToFile (tcFileName,tcSkeleton,tcIgnoreList)
		this._SetSettings()
		
		IF VARTYPE(m.tcFileName)#'C' OR EMPTY(m.tcFileName)
			tcFileName = PUTFILE('','','XML')
		ENDIF

		IF VARTYPE(m.tcSkeleton)#'C' OR EMPTY(m.tcSkeleton)
			tcSkeleton = '*'
		ENDIF

		ADDPROPERTY(_Screen,'cMemoryOutFile',SYS(2023)+JUSTFNAME(m.tcFileName)+'.txt')
		ADDPROPERTY(_Screen,'cMemorySkeleton',m.tcSkeleton)

		IF this._ListMemory('cMemoryList')
			LOCAL laLines[1],lnCount,lnStep,lcName,lcVariableType,lcVal,lnArrayHeight,lnArrayWidth
			LOCAL laEntries[1,5],lnEntryCount,laIgnore[1]

			lnCount = ALINES(m.laLines,_Screen.cMemoryList,4)

			lnEntryCount = 0
			
			IF EMPTY(m.tcIgnoreList)
				m.tcIgnoreList = this.cExclusionList
			ENDIF
			IF !EMPTY(m.tcIgnoreList)
				ALINES(m.laIgnore,UPPER(m.tcIgnoreList),4,',')
			ENDIF

			* fill the entry array from the list memory file
			* any accessible private/public variables will have their values copied directly
			FOR lnStep = 1 TO m.lnCount
				STORE '' TO lcName,lcVariableType,lcVal,lcScope
				STORE 0 TO lnArrayHeight,lnArrayWidth
				this._GetVariableDetails(@m.laLines,m.lnCount,@m.lnStep,@m.lcName,@m.lcVariableType,@m.lcVal,@m.lcScope)
				IF ISUPPER(m.lcVariableType) AND m.lcVariableType#'O'
					IF this.bIgnorePublic AND EMPTY(m.lcScope)
						LOOP
					ENDIF
					IF this.bIgnoreLocal AND m.lcScope='Local'
						LOOP
					ENDIF					
					IF this.bIgnorePrivate AND m.lcScope = 'Priv'
						LOOP
					ENDIF
					IF this.bIgnoreHidden AND m.lcScope = '(hid)'
						LOOP
					ENDIF
				
					IF !EMPTY(m.tcIgnoreList) AND this._IgnoreVariable(@m.laIgnore,@m.lcName)
						LOOP
					ENDIF
				
					IF m.lcScope # 'Local'
						* we can't save objects at the moment (but we could in theory)
						lcVal = this._GetLiveValue(m.lcName,m.lcVariableType,m.lcVal,@m.lnArrayHeight,@m.lnArrayWidth)
					ENDIF

					this._AddEntry(@m.laEntries,@m.lnEntryCount,m.lcName,m.lcVariableType,m.lcVal,m.lnArrayHeight,m.lnArrayWidth)
				ENDIF
			ENDFOR

			IF m.lnEntryCount > 0
				DIMENSION laEntries[m.lnEntryCount,5]
				this._WriteToXml(@m.laEntries,m.lnEntryCount,m.tcFileName)
			ENDIF
		ELSE
			StopHere('List memory file not found.')
		ENDIF && ListMemory

		*!* IF VERS(2)#2 && uncomment to preserve the raw text file
		IF PEMSTATUS(_Screen,'cMemoryOutFile',5) AND FILE(_Screen.cMemoryOutFile)
			DELETE FILE (_Screen.cMemoryOutFile)
		ENDIF
		*!* ENDIF

		REMOVEPROPERTY(_Screen,'cMemoryOutFile')
		REMOVEPROPERTY(_Screen,'cMemorySkeleton')
		REMOVEPROPERTY(_Screen,'cMemoryList')

		this._ClearSettings()
	ENDPROC

	* convert an existing vfp memory file to an xml file
	PROCEDURE ConvertFileToXml (tcFileName)

		ADDPROPERTY(_Screen,'tcFileName',m.tcFileName)

		IF VARTYPE(_Screen.tcFileName)#'C' OR EMPTY(_Screen.tcFileName)
			_Screen.tcFileName = GETFILE()
			IF EMPTY(_Screen.tcFileName)
				RETURN .f.
			ENDIF
		ENDIF

		IF !FILE(_Screen.tcFileName)
			MESSAGEBOX('File '+_Screen.tcFileName+' not found.')
			RETURN .f.
		ENDIF

		this._SetSettings()

		ADDPROPERTY(_Screen,'cMemoryOutFile',SYS(2023)+'memfile.txt')
		ADDPROPERTY(_Screen,'cMemorySkeleton','*')

		IF this._ListMemory('cMemoryListBefore')
			*!* * using "LIKE *" prevents VFP from listing the internal memory variables
			*!* * so we don't need to call this._IsLastLineOfMemory
			*!* LIST MEMORY LIKE * TO FILE 'memfile.txt' NOCONSOLE
			*!* ADDPROPERTY(_Screen,'cMemoryListBefore',FILETOSTR('memfile.txt'))

			LOCAL l__oException
			TRY
				RESTORE FROM (_Screen.tcFileName) ADDITIVE
			CATCH TO l__oException
				MESSAGEBOX('Exception when reading memory file'+CHR(13)+CHR(10)+m.l__oException.Message)
			ENDTRY

			*!* LIST MEMORY LIKE * TO FILE 'memfile.txt' NOCONSOLE
			IF this._ListMemory('cMemoryListAfter')
				*!* ADDPROPERTY(_Screen,'cMemoryListAfter',FILETOSTR('memfile.txt'))

				* compare the before and after to find the variables that have changed
				LOCAL laResult[10,5],lnCount
				lnCount = this._CompareBeforeAndAfter(@m.laResult,_Screen.cMemoryListBefore,_Screen.cMemoryListAfter)

				IF m.lnCount > 0
					DIMENSION laResult[m.lnCount,5]
					this._WriteToXml(@m.laResult,m.lnCount,LOWER(_Screen.tcFileName)+'.xml')
				ENDIF
			ENDIF && ListMemory
		ENDIF && ListMemory

		IF VERS(2)#2
			DELETE FILE ('memfile.txt')
		ENDIF

		REMOVEPROPERTY(_Screen,'cMemoryListBefore')
		REMOVEPROPERTY(_Screen,'cMemoryListAfter')
		REMOVEPROPERTY(_Screen,'tcFileName')

		this._ClearSettings()
	ENDPROC

	* scan through the array and match against the variable name
	PROTECTED PROCEDURE _IgnoreVariable (raIgnoreList,tcName)
		EXTERNAL ARRAY raIgnoreList
		LOCAL lcString
		FOR EACH lcString IN raIgnoreList
			IF IIF(lcString='"',qw(tcName)=lcString,tcName = lcString)
				RETURN .t.
			ENDIF
		ENDFOR

		RETURN .f.
	ENDPROC

	* set a variable to a value stored in the t__cValue parameter
	PROTECTED PROCEDURE _SetVariable (t__cName,t__cType,t__cValue)
		LOCAL loException
		TRY
			IF m.t__cType#'C' AND (INLIST(m.t__cType,'D','T') OR TYPE(m.t__cValue)#'U')
				* for non-character fields we have to evaluate the value
				IF INLIST(m.t__cType,'D','T','O')
					IF m.t__cType = 'O'
						* object types cannot be converted
						t__cValue = .f.
					ELSE
						* dates need to be parsed so the eval works
						t__cValue = this._ParseDate(m.t__cValue,m.t__cType='T')
					ENDIF
				ELSE
					t__cValue = EVAL(m.t__cValue)
				ENDIF
			ENDIF

			* set the variable
			IF this.bObjectMode
				IF '['$m.t__cName
					* it's an array element
					STORE m.t__cValue TO ('this.oMemoryObject.'+m.t__cName)
				ELSE
					* object mode stores the variables to an object
					ADDPROPERTY(this.oMemoryObject,m.t__cName,m.t__cValue)
				ENDIF
			ELSE
				* otherwise store it to a memory variable
				STORE m.t__cValue TO (m.t__cName)
			ENDIF
		CATCH TO loException
			LOCAL l__cMessage
			l__cMessage = 'MemToXml variable read error ('+t__cName+'):'+CHR(13)+CHR(10)+loException.Message
			wait window m.l__cMessage TIMEOUT .2
			whereami(m.l__cMessage)
		ENDTRY
		RETURN

	ENDPROC

	* build the array definition from xml tree structure
	* <var name="arrayName" type="A" height="10" width="2">
	* 	<cell id="1,1" type="N">12</cell>
	* 	<cell id="1,2" type="C"><![CDATA[cell contents]]</cell>
	* 	..
	* </var>
	PROTECTED PROCEDURE _GetArrayFromXml (raElements,tcArrayName,toXmlNode,rnHeight,rnWidth)
		EXTERNAL ARRAY raElements
		
		LOCAL loNode,lnStep,loNameAttr,loTypeAttr,lcValue,lnRow,lnCol,lnAt
		loNode = toXmlNode.firstChild
		IF VARTYPE(loNode)#'O'
			RETURN 0
		ENDIF
		DIMENSION raElements[toXmlNode.childNodes.length,3]
		
		lnStep = 0
		* scan through the cell elements by skipping down the list through the nextSibling
		DO WHILE NOT ISNULL(loNode)
			* get the attributes of the cell element
			loNameAttr = loNode.attributes.getNamedItem('id')
			loTypeAttr = loNode.attributes.getNamedItem('type')
			IF ISNULL(m.loNameAttr) OR ISNULL(m.loTypeAttr)
				StopHere('Invalid xml array definition! '+tcArrayName+CHR(13)+CHR(10)+loNode.xml)
				LOOP
			ENDIF
			

			* get the text from the node
			lcValue = loNode.text
			IF VARTYPE(m.loNode.firstChild)='O' AND loNode.firstChild.nodeName = '#cdata'
				IF NOT CHR(13)+CHR(10)$m.lcValue
					lcValue = STRTRAN(m.lcValue,CHR(10),CHR(13)+CHR(10))
				ENDIF
			ENDIF

			* fixed for SET SEPARATOR TO ',' where VAL() interprets the comma as part of the number
			lnAt = AT(',',loNameAttr.Value)
			IF lnAt > 0
				* the "name" of the cell is it's x,y position in the array
				lnRow = VAL(LEFT(loNameAttr.Value,lnAt-1))
				lnCol = VAL(SUBSTR(loNameAttr.Value,lnAt+1))
			ELSE
				lnRow = VAL(loNameAttr.Value)
				lnCol = 0
			ENDIF
			
			* add a new element to the list
			lnStep = lnStep + 1
			raElements[m.lnStep,1] = m.tcArrayName+'['+TRAN(m.lnRow)+IIF(EMPTY(m.lnCol),'',','+TRAN(m.lnCol))+']'
			raElements[m.lnStep,2] = m.loTypeAttr.Value
			raElements[m.lnStep,3] = m.lcValue
			
			loNode = loNode.nextSibling
			
			* adjust the max height/width of the array (not strictly necessary)
			rnHeight = MAX(m.rnHeight,m.lnRow)
			rnWidth = MAX(m.rnWidth,m.lnCol)
		ENDDO

		
		RETURN lnStep
	ENDPROC
	
	* convert the array string (from LIST MEMORY) into a series of lines that will define the array
	PROTECTED PROCEDURE _GetArrayFromString (raElements,tcArrayName,tcContents,rnHeight,rnWidth)

		EXTERNAL ARRAY raElements

		LOCAL laLines[1],lnCount,lnStep,lnArrayHeight,lnArrayWidth,loException
		lnCount = ALINES(m.laLines,m.tcContents,4) && do not include empty elements
		IF m.lnCount == 0
			Stophere(m.tcArrayName+': Array with no elements')
			RETURN ''
		ENDIF

		LOCAL lnEStep,lcIndex,lcValue,lcVariableType,lnAt,lnCol,lnRow
		STORE 0 TO m.lnEStep

		DIMENSION raElements[m.lnCount,4]
		FOR lnStep = 1 TO m.lnCount
			lnEStep = m.lnEStep + 1
			* get the text from between the parens
			lcIndex = STREXTRACT(m.laLines[m.lnStep],'(',')')

			IF EMPTY(m.lcIndex)
				Stophere(m.tcArrayName+': invalid sub-clause '+m.laLines[m.lnStep])
				LOOP
			ENDIF
			
			* get the type and value based on an offset from the close paren
			lnAt = AT(')',m.laLines[m.lnStep])
			lcVariableType = SUBSTR(m.laLines[m.lnStep],m.lnAt+6,1)
			lcValue = SUBSTR(m.laLines[m.lnStep],m.lnAt+9)
			IF m.lcValue = '"'
				* multiline elements in the array, so try to make sure we respect the new lines
				DO WHILE m.lnStep < m.lnCount AND LTRIM(m.laLines[m.lnStep+1])#'(' OR RIGHT(m.lcValue,1)#'"'
					lnStep = m.lnStep + 1
					lcValue = m.lcValue + CHR(13)+CHR(10) + m.laLines[m.lnStep]
				ENDDO
			ENDIF

			* add a new element to the array
			* determine its position in the array
			lnAt = AT(',',m.lcIndex) && fix for SET SEPARATOR TO ','
			IF lnAt > 0
				lnRow = VAL(LEFT(m.lcIndex,m.lnAt-1))
				lnCol = VAL(SUBSTR(m.lcIndex,m.lnAt+1))
			ELSE
				lnRow = VAL(m.lcIndex)
				lnCol = 0
			ENDIF
			* set the array element's name
			raElements[m.lnEStep,1] = m.tcArrayName+'['+TRAN(m.lnRow)+IIF(EMPTY(m.lnCol),'',','+TRAN(m.lnCol))+']'
			IF TYPE(raElements[m.lnEStep,1])#'U'
				* try to get the live value of the array from memory
				TRY
					lcValue = TRAN(EVAL(raElements[m.lnEStep,1]))
					lcVariableType = TYPE(raElements[m.lnEStep,1])
				CATCH TO loException
					* oops, we can use the value from the LIST MEMORY file
					LOCAL lcMessage
					lcMessage = 'MemToXml array variable read error ('+raElements[m.lnEStep,1]+'):'+CHR(13)+CHR(10)+loException.Message
					*!* wait window m.lcMessage TIMEOUT .2
					whereami(m.lcMessage)
				ENDTRY
			ENDIF
			* set the array element's type
			raElements[m.lnEStep,2] = m.lcVariableType
			* set the array element's value
			raElements[m.lnEStep,3] = m.lcValue

			* adjust the max height/width of the array (not strictly necessary)
			rnHeight = MAX(m.rnHeight,m.lnRow)
			rnWidth = MAX(m.rnWidth,m.lnCol)

		ENDFOR
		RETURN m.lnEStep
	ENDPROC

	PROTECTED PROCEDURE _ParseDate (tcString,tbDateTime)
		LOCAL lcYear,lcMonth,lcDay,lcOutput,ldOutput

		lcOutput = IIF(m.tbDateTime,'DATETIME(','DATE(')

		lcYear = LEFT(m.tcString,4)
		lcMonth = SUBSTR(m.tcString,6,2)
		lcDay = SUBSTR(m.tcString,9,2)

		IF EMPTY(VAL(m.lcYear))
			ldOutput = IIF(tbDateTime,{/:},{})
			*!* lcOutput = IIF(m.tbDateTime,'{/:}','{}')
		ELSE
			*!* StopHere(lcYear+' '+lcMonth+' '+lcDay)
			LOCAL lcHour,lcMinute,lcSecond
			IF m.tbDateTime
				lcHour = SUBSTR(m.tcString,12,2)
				lcMinute = SUBSTR(m.tcString,15,2)
				lcSecond = SUBSTR(m.tcString,18,2)
			ENDIF
			ldOutput = IIF(tbDateTime,DATETIME(VAL(lcYear),VAL(lcMonth),VAL(lcDay),VAL(lcHour),VAL(lcMinute),VAL(lcSecond)),DATE(VAL(lcYear),VAL(lcMonth),VAL(lcDay)))
			*!* lcOutput = m.lcOutput + m.lcYear+','+m.lcMonth+','+m.lcDay

			*!* IF m.tbDateTime
				*!* LOCAL lcHour,lcMinute,lcSecond
				*!* lcHour = SUBSTR(m.tcString,12,2)
				*!* lcMinute = SUBSTR(m.tcString,15,2)
				*!* lcSecond = SUBSTR(m.tcString,18,2)

				*!* lcOutput = m.lcOutput + ','+m.lcHour+','+m.lcMinute+','+m.lcSecond
			*!* ENDIF

			*!* lcOutput = m.lcOutput + ')'
		ENDIF

		RETURN ldOutput
		*!* RETURN m.lcOutput
	ENDPROC

	PROTECTED PROCEDURE _GetDomDoc
		LOCAL loDocument
		IF VARTYPE(m.loDocument)#'O'
			TRY
				loDocument=CREATEOBJECT('msxml2.domdocument.6.0')
			CATCH
			ENDTRY
			IF VARTYPE(m.loDocument)#'O'
				TRY
					loDocument=CREATEOBJECT('msxml2.domdocument') && instantiates DOMDocument.3.0
				CATCH
				ENDTRY
				if vartype(m.loDocument)#'O'
					MESsAGEBOX('MSXML Parser not available on this machine.')
					return .f.
				endif
			ENDIF
		ENDIF

		RETURN m.loDocument
	ENDPROC

	PROTECTED PROCEDURE _ListMemory (tcScreenProperty)
		IF VARTYPE(_Screen.cMemorySkeleton)='C' AND VARTYPE(_Screen.cMemoryOutFile)='C'
			LOCAL lcOldAlt
			* try to prevent the output of the LIST MEMORY from going to the printer (for example in isolate.prg)
			lcOldAlt = SET('Alternate')
			SET ALTERNATE OFF
		
			* read the memory file before and after restoring the memory
			LIST MEMORY LIKE (_Screen.cMemorySkeleton) TO FILE (_Screen.cMemoryOutFile) NOCONSOLE 
			
			IF lcOldAlt = 'ON'
				SET ALTERNATE ON
			ENDIF
			
			IF !EMPTY(tcScreenProperty)
				ADDPROPERTY(_Screen,tcScreenProperty,FILETOSTR(_Screen.cMemoryOutFile))
			ENDIF
			
			RETURN .t.
		ENDIF
		RETURN .f.
	ENDPROC

	PROTECTED PROCEDURE _SetSettings ()
		ADDPROPERTY(_Screen,'oMemSettings',CREATEOBJECT('EMPTY'))

		ADDPROPERTY(_Screen.oMemSettings,'OldSafety',SET("SAFETY"))
		SET SAFETY OFF
		ADDPROPERTY(_Screen.oMemSettings,'OldDate',SET("DATE"))
		SET DATE YMD
		ADDPROPERTY(_Screen.oMemSettings,'OldCentury',SET("CENTURY"))
		SET CENTURY ON
		ADDPROPERTY(_Screen.oMemSettings,'OldHours',SET("HOURS"))
		SET HOURS TO 24
		ADDPROPERTY(_Screen.oMemSettings,'OldPoint',SET("POINT"))
		SET POINT TO '.'
		ADDPROPERTY(_Screen.oMemSettings,'OldSeparator',SET("SEPARATOR"))
		SET SEPARATOR TO ','
	ENDPROC

	PROTECTED PROCEDURE _ClearSettings
		IF VARTYPE(_Screen.oMemSettings)='O'
			IF _Screen.oMemSettings.OldSafety='ON'
				SET SAFETY ON
			ENDIF
			IF _Screen.oMemSettings.OldCentury = 'OFF'
				SET CENTURY OFF
			ENDIF
			SET DATE (_Screen.oMemSettings.OldDate)
			SET HOURS TO (_Screen.oMemSettings.OldHours)
			SET POINT TO (_Screen.oMemSettings.OldPoint)
			SET SEPARATOR TO (_Screen.oMemSettings.OldSeparator)
			REMOVEPROPERTY(_Screen,'oMemSettings')
		ENDIF
	ENDPROC

	PROTECTED PROCEDURE _WriteToXml (raResult,tnCount,tcFileName)
		EXTERNAL ARRAY raResult
		LOCAL lnStep

		IF FILE(FULLPATH(m.tcFileName))
			DELETE FILE (FULLPATH(m.tcFileName))
		ENDIF

		*!* LOCAL lcOldCentury,lcOldDate,lnOldHours
		*!* lcOldCentury = SET('CENTURY')
		*!* SET CENTURY ON
		*!* lcOldDate = SET('DATE')
		*!* SET DATE YMD
		*!* lnOldHours = SET('HOURS')
		*!* SET HOURS TO 24

		SET TEXTMERGE TO (m.tcFileName) NOSHOW
		SET TEXTMERGE DELIMITERS TO '{{','}}'
		SET TEXTMERGE ON

		\\<?xml version = "1.0" encoding="Windows-1252" standalone="yes"?>

		\<memory name="{{STRTRAN(JUSTSTEM(m.tcFileName),'&','&amp;')}}">

		IF this.bSortVariables
			ASORT(m.raResult,1)
		ENDIF

		FOR lnStep = 1 TO m.tnCount
			lcName = m.raResult[m.lnStep,1]
			lcVariableType = m.raResult[m.lnStep,2]
			lcVal = m.raResult[m.lnStep,3]

			DO CASE
				CASE m.lcVariableType = 'N'
					lcVal = TRANSFORM(VAL(m.lcVal))
				CASE m.lcVariableType = 'A'
					this._ArrayToXml(m.lcName,m.lcVal,m.raResult[m.lnStep,4],m.raResult[m.lnStep,5])
					LOOP
				CASE m.lcVariableType = 'C'
					lcVal = '<![CDATA['+m.lcVal+']]>'
			ENDCASE

			\{{CHR(9)}}<var name="{{m.lcName}}" type="{{m.lcVariableType}}">{{m.lcVal}}</var>
		ENDFOR
		
		\</memory>

		SET TEXTMERGE OFF
		SET TEXTMERGE TO
		SET TEXTMERGE DELIMITERS

		*!* IF m.lcOldCentury = 'OFF'
			*!* SET CENTURY OFF
		*!* ENDIF
		*!* SET DATE (m.lcOldDate)
		*!* SET HOURS TO (m.lnOldHours)
	ENDPROC

	* write out the information for the array in xml style
	PROTECTED PROCEDURE _ArrayToXml (tcName, tcContents, tnHeight, tnWidth)
		LOCAL laElements[1],lnCount,lnStep,lcValue,lcVariableType,lcName,lbNoContent
		
		* create the initial element (with the field name, type, array height and width)
		\{{CHR(9)}}<var name="{{m.tcName}}" type="A" height="{{m.tnHeight}}" width="{{m.tnWidth}}">
		
		* since we're using the LIST MEMORY string as the structure, we need to parse
		* that into an actual array that we can use
		* the elements in the resulting array are:
		* 	[1] - the name (eg. arrayName[1,2])
		* 	[2] - the type (eg. N)
		* 	[3] - the value (eg. 12)
		lnCount = this._GetArrayFromString(@laElements,m.tcName,m.tcContents,m.tnHeight,m.tnWidth)
		* now we need to create an xml element for each cell in the array
		FOR lnStep = 1 TO lnCount
			* get the properties of each element
			* for a name we use the x,y coordinate
			lcName = STREXTRACT(m.laElements[m.lnStep,1],'[',']')
			lcVariableType = m.laElements[m.lnStep,2]
			lcValue = m.laElements[m.lnStep,3]
			lbNoContent = .f.
			DO CASE 
				CASE m.lcVariableType='N'
					lcValue = TRAN(VAL(m.lcValue))
				CASE m.lcVariableType='C'
					IF LEN(m.lcValue)=0
						lbNoContent = .t.
					ELSE
						* surround strings with cdata so they can't screw up the xml (though they probably still could)
						lcValue = '<![CDATA['+m.lcValue+']]>'
					ENDIF
				CASE m.lcVariableType='L'
					* the default value for the array is .f. so there's no need to encode these entries
					IF UPPER(m.lcValue)='.F.' AND NOT EMPTY(m.lcName)
						LOOP
					ENDIF
			ENDCASE
			\{{CHR(9)}}{{CHR(9)}}<cell id="{{m.lcName}}" type="{{m.lcVariableType}}"
			IF m.lbNoContent
				\\/>
			ELSE
				\\>{{m.lcValue}}</cell>
			ENDIF
		ENDFOR
		
		\{{CHR(9)}}</var>
	ENDPROC
	
	* for converting from the foxpro memory files to the xml
	* this is not really used in the live program (though it's included for utility purposes)
	* see SaveToFile/RestoreToFile instead
	PROTECTED PROCEDURE _CompareBeforeAndAfter(raResult,lcBefore,lcAfter)
		EXTERNAL ARRAY raResult
		LOCAL laBefore[1],laAfter[1],lnBeforeCount,laAfterCount,lnBeforeStep,lnAfterStep,lnNewEntries,lcAfterScope

		lnBeforeCount = ALINES(m.laBefore,m.lcBefore,4)
		lnAfterCount = ALINES(m.laAfter,m.lcAfter,4)

		LOCAL lcBeforeName,lcBeforeType,lcBeforeVal,lbBeforeOk,laIgnore[1],lnIgnoreCount
		LOCAL lcAfterName,lcAfterType,lcAfterVal,lbAfterOk,lnArrayHeight,lnArrayWidth,lbAdvanceBefore

		lbAdvanceBefore = .t.
		lnIgnoreCount = 0
		IF !EMPTY(this.cExclusionList)
			lnIgnoreCount = ALINES(m.laIgnore,UPPER(this.cExclusionList),4,',')
		ENDIF
		lnNewEntries = 0
		
		STORE 0 TO lnBeforeStep,lnAfterStep
		DO WHILE .t.
			IF lbAdvanceBefore
				STORE .f. TO lbBeforeOk
				STORE '' TO lcBeforeName,lcBeforeType
				* find the next non-empty line
				FOR lnBeforeStep = m.lnBeforeStep+1 TO m.lnBeforeCount
					*!* IF this._IsLastLineOfMemory(laBefore[lnBeforeStep])
					*!* lbBeforeOk = .f.
					*!* EXIT
					*!* ENDIF

					IF laBefore[m.lnBeforeStep]#' '
						lbBeforeOk = .t.
						lbAdvanceBefore = .f.
						EXIT
					ENDIF
				ENDFOR
			ENDIF

			STORE .f. TO lbAfterOk
			* find the next non-empty line
			FOR lnAfterStep = m.lnAfterStep+1 TO m.lnAfterCount
				*!* IF this._IsLastLineOfMemory(laAfter[lnAfterStep])
				*!* lbAfterOk = .f.
				*!* EXIT
				*!* ENDIF

				IF laAfter[m.lnAfterStep]#' '
					lbAfterOk = .t.
					EXIT
				ENDIF
			ENDFOR
			
			IF m.lbBeforeOk AND EMPTY(lcBeforeName)
				* get variable information from before the memory was restored
				this._GetVariableDetails(@m.laBefore,m.lnBeforeCount,@m.lnBeforeStep,@m.lcBeforeName,@m.lcBeforeType,@m.lcBeforeVal)
			ENDIF

			IF m.lbAfterOk
				* get variable information from after the memory was restored
				this._GetVariableDetails(@m.laAfter,m.lnAfterCount,@m.lnAfterStep,@m.lcAfterName,@m.lcAfterType,@m.lcAfterVal,@m.lcAfterScope)
				STORE 0 TO lnArrayHeight,lnArrayWidth
				IF m.lcAfterScope # 'Local'
					* try to get the actual value of the variable in memory (rather than relying on the LIST MEMORY entry)
					lcAfterVal = this._GetLiveValue(m.lcAfterName,m.lcAfterType,m.lcAfterVal,@m.lnArrayHeight,@m.lnArrayWidth)
					IF lcAfterType='C'
						lcAfterVal = '"'+lcAfterVal+'"'
					ENDIF
				ENDIF
			ENDIF
			
			IF NOT m.lbAfterOk
				* this means we're out of entries from the memory file, so we're done
				EXIT
			ENDIF
			
			* check if the memory file added/changed this variable
			IF NOT (m.lcBeforeName == m.lcAfterName AND m.lcBeforeType == m.lcAfterType AND (m.lcBeforeVal == m.lcAfterVal OR (lcBeforeType='N' AND VAL(lcBeforeVal)=VAL(lcAfterVal))) )
				* check if the variable is on the excluded list
				IF ISUPPER(m.lcAfterType) AND (m.lnIgnoreCount = 0 OR NOT this._IgnoreVariable(@m.laIgnore,m.lcAfterName))
					* new entry in the list
					this._AddEntry(@m.raResult,@m.lnNewEntries,m.lcAfterName,m.lcAfterType,m.lcAfterVal,m.lnArrayHeight,m.lnArrayWidth)
				ENDIF
			ENDIF

			*!* * advance to the next entry in the memory file list
			*!* lnAfterStep = m.lnAfterStep + 1			
			
			* only advance to the next entry in the original list if we've found a match in the memore file list
			lbAdvanceBefore = (m.lcBeforeName == m.lcAfterName)
			
		ENDDO

		RETURN m.lnNewEntries
	ENDPROC

	PROTECTED PROCEDURE _GetLiveValue (t__cName,t__cType,t__cValue,r__nHeight,r__nWidth)

		IF INLIST(m.t__cName,'L__','T__')
			RETURN m.t__cValue
		ENDIF

		IF m.t__cType = 'O'
			RETURN '.f.'
		ENDIF

		LOCAL l__cResult,l__cActualType
		IF m.t__cType = 'C'
			* strip the quotes
			l__cResult = SUBSTR(m.t__cValue,2,LEN(m.t__cValue)-2)
		ELSE
			l__cResult = m.t__cValue
		ENDIF
		TRY
			IF m.t__cType = 'A'
				IF TYPE(m.t__cName,1)='A'
					LOCAL l__row,l__col,l__cElement,l__cShowValue,l__xEvalValue
					r__nHeight = ALEN(&t__cName,1)
					r__nWidth = ALEN(&t__cName,2)
					l__cResult = ''

					* emulate the LIST MEMORY style with the array
					FOR l__row = 1 TO m.r__nHeight
						FOR l__col = 1 TO EVL(m.r__nWidth,1)
							l__cResult = m.l__cResult + CHR(13)+CHR(10)+SPACE(IIF(m.r__nWidth==0,8,3))+'('+STR(m.l__row,4)+IIF(m.r__nWidth==0,'',','+STR(m.l__col,4))+')'
							l__cElement = m.t__cName+'['+TRAN(m.l__row)+IIF(m.r__nWidth==0,'',','+TRAN(m.l__col))+']'
							l__xEvalValue = EVAL(m.l__cElement)
							l__cActualType = VARTYPE(m.l__xEvalValue)
							DO CASE	
								CASE m.l__cActualType = 'C'
									l__cShowValue = '"'+l__xEvalValue+'"'
								CASE m.l__cActualType = 'N'
									LOCAL l__nActualValue
									l__nActualValue = l__xEvalValue
									l__cShowValue = PADR(TRAN(m.l__nActualValue),12)+'('+STR(m.l__nActualValue,19,8)+')'
								OTHER
									l__cShowValue = TRAN(l__xEvalValue)
							ENDCASE
							l__cResult = m.l__cResult + SPACE(5) + TYPE(m.l__cElement) + SPACE(2) + m.l__cShowValue
						ENDFOR
					ENDFOR
				*!* ELSE
					*!* Stophere('Undefined array '+t__cName)
				ENDIF
			ELSE
				IF TYPE(m.t__cName)#'U' AND TYPE(m.t__cName)=m.t__cType
					IF m.t__cType == 'C'
						* the xml spec apparently doesn't support CRLF -- so convert to the LF character when encoding and convert back when decoding the xml
						l__cResult = EVAL(m.t__cName)
					ELSE
						l__cResult = TRAN(EVAL(m.t__cName))
					ENDIF
				*!* ELSE
					*!* Stophere('Undefined variable '+t__cName)
				ENDIF
			ENDIF
		CATCH TO l__oException
			LOCAL l__cErrMess
			l__cErrMess = 'MemToXml save variable error ('+m.t__cName+'):'+CHR(13)+CHR(10)+m.l__oException.Message
			
			whereami(m.l__cErrMess)
			WAIT WINDOW m.l__cErrMess TIMEOUT .2
		ENDTRY
		RETURN m.l__cResult
	ENDPROC

	* this is no longer useful
	PROTECTED PROCEDURE _IsLastLineOfMemory (tcLine)
		RETURN VAL(m.tcLine)#0 AND 'variables defined,'$m.tcLine AND RIGHT(m.tcLine,10)=='bytes used'
	ENDPROC

	* add an entry to the results array
	PROTECTED PROCEDURE _AddEntry (raArray,rnEntries,tcName,tcType,tcVal,tnHeight,tnWidth)
		EXTERNAL ARRAY raArray
		IF INLIST(m.tcName,'L__','T__')
			RETURN
		ENDIF

		rnEntries = m.rnEntries + 1
		IF ALEN(m.raArray,1) < m.rnEntries
			DIMENSION raArray[m.rnEntries+10,5]
		ENDIF

		raArray[m.rnEntries,1] = m.tcName
		raArray[m.rnEntries,2] = m.tcType
		raArray[m.rnEntries,3] = m.tcVal
		raArray[m.rnEntries,4] = m.tnHeight
		raArray[m.rnEntries,5] = m.tnWidth
	ENDPROC

	* get the variable details from the LIST MEMORY text file
	PROTECTED PROCEDURE _GetVariableDetails (raArray,tnCount,rnStep,rcName,rcType,rcValue,rcScope,rcFunction)
		EXTERNAL ARRAY raArray
		LOCAL lbInQuote,lcFirstLine,lcNextLine,lnFirstWord
		
		lcFirstLine = m.raArray[m.rnStep]

		rcFunction = ''
		rcName = GETWORDNUM(m.lcFirstLine,1)
		lnFirstWord = 2
		IF LEN(m.rcName) > 10
			* the variable's details are listed on the next line
			rnStep = m.rnStep + 1
			lcFirstLine = m.raArray[m.rnStep]
			lnFirstWord = 1
		ENDIF
		
		rcScope = GETWORDNUM(lcFirstLine,lnFirstWord)
		rcType = GETWORDNUM(lcFirstLine,lnFirstWord+1)
		
		*!* rcScope = TRIM(SUBSTR(m.lcFirstLine,13,5))
		
		*!* rcType = SUBSTR(m.lcFirstLine,20,1)
		
		* need to ensure that this value is actually the type rather than a line like this:
		* TCMESSAGE   Priv   lcerror
		IF ISLOWER(rcType)
			* store the name of the value variable
			* but the fact that rcType is LOWERCASE will mean that this entry is skipped
			*!* rcType = SUBSTR(m.lcFirstLine,20)
			RETURN 
		ENDIF
		
		IF LEN(rcType)=1
			rcValue = ''
			
			IF m.rcType # 'A'
				rcValue = LTRIM(SUBSTR(m.raArray[m.rnStep],AT(' '+rcType+' ',m.raArray[m.rnStep])+2))
				IF INLIST(rcScope,'Priv','Local','(hid)') 
					rcFunction = this._StripFunctionName(@rcValue,IIF(m.rcType='C','"',''))
				ENDIF
			ENDIF
			
			*!* IF rcName = 'SYS_ACOSTING'
				*!* StopHere(m.rcName)
			*!* ENDIF
			
			lbInQuote = m.rcType = 'C' AND (LEN(m.rcValue)==1 OR RIGHT(m.rcValue,1)#'"') AND NOT ('...'$rcValue AND RIGHT(rcValue,5)='bytes')
			* handle the case where there are multiple lines for one variable (an array or a string)
			DO WHILE m.rnStep < m.tnCount AND (m.lbInQuote OR m.raArray[m.rnStep+1]=' ') && AND !this._IsLastLineOfMemory(raArray[rnStep+1])
				rnStep = m.rnStep + 1
				lcNextLine = m.raArray[m.rnStep]
				rcValue = m.rcValue + IIF(!EMPTY(LEFT(m.lcNextLine,20)),CHR(13)+CHR(10)+m.lcNextLine,SUBSTR(m.lcNextLine,23))
				IF m.rcType = 'C' AND INLIST(m.rcScope,'Priv','Local','(hid)')
					rcFunction = this._StripFunctionName(@m.rcValue,'"')
				ENDIF
				m.lbInQuote = (m.lbInQuote OR SUBSTR(m.lcNextLine,23)='"') AND RIGHT(m.rcValue,1)#'"' AND NOT ('...'$rcValue AND RIGHT(rcValue,5)='bytes')
				IF !m.lbInQuote AND rcType='A'
					* check if the array element is listed as a multi-line character string
					IF lcNextLine=' ' AND LEFT(LTRIM(lcNextLine),1)='(' AND ('"'$m.lcNextLine AND RIGHT(m.lcNextLine,1)#'"'  AND NOT ('...'$lcNextLine AND RIGHT(lcNextLine,5)='bytes'))
						lbInQuote = .t.
					ENDIF
				ENDIF
			ENDDO			
		ENDIF

	ENDPROC

	* VFP adds the declaring function to the end of the value (we don't want it!)
	* for strings we can't just match against two spaces and a word..
	* .. we want to find '"  <function_name>', so pass the '"' character in as the prefix
	PROTECTED PROCEDURE _StripFunctionName (rcString,tcPrefix)
		LOCAL lnSpaceAt,lcFunctionName
		* find if there's a suffix
		lnSpaceAt = RAT(tcPrefix+'  ',m.rcString)
		* make sure that the suffix is one word only (otherwise it's just a random string that matches)
		IF m.lnSpaceAt > 0
			lcFunctionName = SUBSTR(m.rcString,m.lnSpaceAt+2+LEN(m.tcPrefix))
			IF GETWORDCOUNT(m.lcFunctionName) = 1
				rcString = LEFT(m.rcString,m.lnSpaceAt-1+LEN(tcPrefix))
				RETURN m.lcFunctionName
			ENDIF
		ENDIF	
	ENDPROC

ENDDEFINE


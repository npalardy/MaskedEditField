#tag Class
Protected Class MaskedField
Inherits TextField
	#tag Event
		Function KeyDown(Key As String) As Boolean
		  Dim newtext As String
		  dim befor as string
		  dim after as string
		  dim i as integer
		  dim b as boolean
		  dim sEnd As integer
		  dim curSelStart as integer
		  dim curSelLen as integer
		  
		  'generally let the user decide about accepting editing keys
		  if asc(key) <= 31 then
		    return KeyDown(key)
		  end if
		  
		  ' if we have no mask accept what the user says in their code
		  if ubound(myMask) < 0 then
		    return KeyDown(key)
		  end if
		  
		  curSelStart = me.SelStart
		  curSelLen = me.SelLength
		  
		  if uBound( myMask ) > -1 then ' We prevent insertion that may violate the mask.
		    
		    ' begin added P2 030428
		    ' This block checks the selection in the text & sets t.
		    sEnd = me.selStart + me.selLength
		    if sEnd = len( me.text ) then ' Cursor is at end of text or selects to end of text.
		      ' Set to to the portion of the text before the selection or all for a selection point at the end.
		      befor = left( me.text, me.selStart )
		      after = ""
		    else ' Cursor is positioned to insert in text
		      
		      ' rejected()
		      ' return true
		      befor = left ( me.text, me.selStart )
		      after = mid ( me.text, me.selStart + 1 )
		    end if
		    ' end added P2 030428
		    
		  end if
		  
		  for i = 0 to ubound(myMask)
		    
		    ' if this key makes the text still match ONE of the masks then it's OK
		    
		    'if matches(me.text + key,  myMask(i) ) then ' replaced P2 030428 with ...
		    ' if matches( t + key,  myMask(i) ) then ' replaced njp 030429
		    newtext = befor + key + after
		    
		    if matches( newtext ,  myMask(i) ) then
		      
		      ' call the users key function
		      b = KeyDown(key)
		      
		      me.text = newtext
		      
		      me.selstart = curSelStart + 1
		      me.sellength = curSelLen
		      
		      ' key is OK
		      return true
		    End if
		    
		  next
		  
		  ' call the users key function
		  b = KeyDown(key)
		  
		  rejected key
		  
		  me.selstart = curSelStart + 1
		  me.sellength = curSelLen
		  
		  ' key is not OK
		  return true
		  
		End Function
	#tag EndEvent

	#tag Event
		Sub LostFocus()
		  dim i as integer
		  dim newtext as string
		  
		  for i = 0 to ubound(myMask)
		    
		    newtext = me.text
		    if matches( newtext ,  myMask(i) ) then
		      
		      me.text = newtext
		      
		    end if
		    
		  next
		  
		  LostFocus
		End Sub
	#tag EndEvent

	#tag Event
		Sub TextChange()
		  dim i as integer
		  dim j as integer
		  dim newText as string
		  dim oksofar as boolean
		  
		  for j = 1 to len(me.text)
		    newText = mid(me.text,1, j)
		    
		    okSoFar = false
		    
		    for i = 0 to ubound(myMask)
		      
		      ' if this text matches ONE of the masks then it's OK
		      if matches(newtext,  myMask(i) ) then 
		        okSoFar = true
		      End if
		      
		    next
		    
		    if oksofar = false then
		      me.text = mid(me.text,1, j-1)
		      me.selstart = len(me.text)
		      return
		    end if
		    
		  next
		  
		End Sub
	#tag EndEvent


	#tag Method, Flags = &h0
		Sub clearMasks()
		  redim myMask(-1)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub clearMessage()
		  myMessage = ""
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function isAlpha(char as string) As boolean
		  
		  return instr(1,"abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ", char) > 0 
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function isAlphaNum(char as string) As boolean
		  return instr(1,"abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789", char) > 0 
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function isAscii(char as string) As boolean
		  return asc(char) < 128 
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function isDigit(char as string) As boolean
		  return instr(1,"0123456789", char) > 0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function isNotAlpha(char as string) As boolean
		  return instr(1,"abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ", char) = 0 
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function isNotAlphaNum(char as string) As boolean
		  return instr(1,"abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789", char) = 0 
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function isNotDigit(char as string) As boolean
		  return instr(1,"0123456789", char) = 0
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function matches(byref text as string , pattern as string) As boolean
		  
		  dim char as string
		  dim pchar as string
		  dim pPos as integer
		  'dim matchExact as boolean - removed P2 030428, appears redundant.
		  dim i as integer
		  dim temp as string
		  
		  ' # is a digit
		  ' $ is a non-digit - added P2 030428
		  ' A is a char A - Z
		  ' l (small ell) is a char (forced to lower case)
		  ' L is a char (forced to uppercase)
		  ' ! is a non alpha - added P2 030428
		  ' . is any char
		  ' \x is a x (must be used for #, $, A, l, L, !, ., \)
		  ' matches everything else literally
		  
		  pPos = 1 
		  for i = 1 to len(text)
		    
		    char = mid(text,i,1)
		    
		    if i > len(pattern) then 
		      return false
		    end if
		    
		    'matchExact = false
		    
		    pchar = mid(pattern,pPos,1)
		    
		    if pchar = "^" then ' the NOT flag
		      pPos = pPos + 1
		      pchar = pchar + mid(pattern,pPos,1)
		    end if
		    
		    if pchar = "\" then
		      'matchExact = true
		      pPos = pPos + 1
		      pchar = mid(pattern,pPos,1)
		      if pchar <> char then
		        return false
		      else
		        temp = temp + char
		      end if
		    else
		      if strcomp(pchar,"A",0)=0 and (isAlpha(char)) then
		        temp = temp + char
		      elseIf strcomp(pchar ,"^A",0)=0 and isNotAlpha( char ) then ' added P2 030428
		        temp = temp + char
		      elseif strcomp(pchar ,"#",0) = 0 and (isDigit(char)) then
		        temp = temp + char
		      elseif strcomp(pchar ,"^#",0) = 0 and isNotDigit( char ) then ' added P2 030428
		        temp = temp + char
		      elseif strcomp(pchar, "*",0) = 0 and isAlphaNum( char ) then ' added P2 030429
		        temp = temp + char
		      elseif strcomp(pchar ,"^*" ,0) = 0and isNotAlphaNum( char ) then ' added P2 030429
		        temp = temp + char
		      elseif strcomp(pchar,"L",0) = 0 and (isalpha(char)) then
		        temp = temp + uppercase(char)      
		      elseif strcomp(pchar,"^L",0) = 0 and (isNotalpha(char)) then
		        temp = temp + char
		      elseif strcomp(pchar, "l",0) = 0 and (isalpha(char)) then
		        temp = temp + lowercase(char)
		      elseif strcomp(pchar, "^l",0) = 0 and (isNotalpha(char)) then
		        temp = temp + char
		      elseif strcomp(pchar, char,0) = 0 then
		        temp = temp + char
		      elseif strcomp(pchar, ".",0) = 0 then
		        temp = temp + char
		      else
		        return false
		      end if
		    end if
		    
		    pPos = pPos + 1
		  next
		  
		  text = temp
		  return true
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub setMask(maskString as string)
		  
		  myMask.append maskString
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub setMessage(message As string)
		  myMessage = message
		End Sub
	#tag EndMethod


	#tag Hook, Flags = &h0
		Event KeyDown(key as string) As boolean
	#tag EndHook

	#tag Hook, Flags = &h0
		Event LostFocus()
	#tag EndHook

	#tag Hook, Flags = &h0
		Event Rejected(key as string)
	#tag EndHook


	#tag Note, Name = Documentation
		Implements a simple masked edit field
		Each field can have multiple masks set for it
		
		You set the mask and the masked field handles everything else
		
		A mask is specified as :
		
		          a # character in the mask signifies a number (0-9)
		          a L character in the mask signifies a letter (will fold to upper case) // njp 030430
		          a l (small ell) character in the mask signifies a letter (will fold to lower case) // njp 030430
		          a A character in the mask signifies a letter (upper or lower case)
		          a . character in the mask signifies any character
		          a * character in the mask signifies any letter or number          // added P2 030429
		
		          any other character (except the \) is required as is
		
		          a ^ character preceding any other character signifies NOT the character that follows
		          
		          e.g.  ^# signifies any character that is not a number
		
		          a \ character precedes any character that requires escaping (only a #, A, *, ^, . and \)
		          
		special notes :
		          ^A is NOT an alpha character
		          ^L is NOT an alpha character
		          ^l is NOT an alpha character
		          all 3 cases have similar semantics and will allow the SAME set of characters
		          
		          ^. has no defined meaning and causes input to STOP being accepted at the point it appears in the mask
		          
		          ^* means anything NOT a number or letter
		          
		METHODS
		===============================================================
		clearMasks
		   clears all masks from the edit field
		   
		setMask
		   adds a mask to the edit field
		   
		
		EVENTS
		===============================================================
		KeyDown(key as string)
		     this event will be propagated to the user of the class
		     but, returning true/false is not useful and will not behave as it would for a regular edit field
		     
		LostFocus
		     will do any character folding that is required by the mask that is set
		     this event will be propagated to the user of the class
		     
		Rejected(key as string)
		    the key that has been rejected as it causes masks to not match
		
		
	#tag EndNote


	#tag Property, Flags = &h1
		Protected myMask(-1) As string
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected myMessage As string
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="AllowAutoDeactivate"
			Visible=true
			Group="Appearance"
			InitialValue="True"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="BackgroundColor"
			Visible=true
			Group="Appearance"
			InitialValue="&hFFFFFF"
			Type="Color"
			EditorType="Color"
		#tag EndViewProperty
		#tag ViewProperty
			Name="HasBorder"
			Visible=true
			Group="Appearance"
			InitialValue="True"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Tooltip"
			Visible=true
			Group="Appearance"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Transparent"
			Visible=true
			Group="Appearance"
			InitialValue="False"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="AllowFocusRing"
			Visible=true
			Group="Appearance"
			InitialValue="True"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="FontName"
			Visible=true
			Group="Font"
			InitialValue="System"
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="FontSize"
			Visible=true
			Group="Font"
			InitialValue="0"
			Type="Single"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="FontUnit"
			Visible=true
			Group="Font"
			InitialValue="0"
			Type="FontUnits"
			EditorType="Enum"
			#tag EnumValues
				"0 - Default"
				"1 - Pixel"
				"2 - Point"
				"3 - Inch"
				"4 - Millimeter"
			#tag EndEnumValues
		#tag EndViewProperty
		#tag ViewProperty
			Name="Hint"
			Visible=true
			Group="Initial State"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="AllowTabs"
			Visible=true
			Group="Behavior"
			InitialValue="False"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="TextAlignment"
			Visible=true
			Group="Behavior"
			InitialValue="0"
			Type="TextAlignments"
			EditorType="Enum"
			#tag EnumValues
				"0 - Default"
				"1 - Left"
				"2 - Center"
				"3 - Right"
			#tag EndEnumValues
		#tag EndViewProperty
		#tag ViewProperty
			Name="AllowSpellChecking"
			Visible=true
			Group="Behavior"
			InitialValue="False"
			Type="boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="MaximumCharactersAllowed"
			Visible=true
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="ValidationMask"
			Visible=true
			Group="Behavior"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue=""
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Width"
			Visible=true
			Group="Position"
			InitialValue="80"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Height"
			Visible=true
			Group="Position"
			InitialValue="22"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="LockLeft"
			Visible=true
			Group="Position"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="LockTop"
			Visible=true
			Group="Position"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="LockRight"
			Visible=true
			Group="Position"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="LockBottom"
			Visible=true
			Group="Position"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="TabPanelIndex"
			Visible=false
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="TabIndex"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="TabStop"
			Visible=true
			Group="Position"
			InitialValue="True"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Password"
			Visible=true
			Group="Appearance"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="TextColor"
			Visible=true
			Group="Appearance"
			InitialValue="&h000000"
			Type="Color"
			EditorType="Color"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Enabled"
			Visible=true
			Group="Appearance"
			InitialValue="True"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Format"
			Visible=true
			Group="Appearance"
			InitialValue=""
			Type="String"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Visible"
			Visible=true
			Group="Appearance"
			InitialValue="True"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Bold"
			Visible=true
			Group="Font"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Italic"
			Visible=true
			Group="Font"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Underline"
			Visible=true
			Group="Font"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="Text"
			Visible=true
			Group="Initial State"
			InitialValue=""
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ReadOnly"
			Visible=true
			Group="Behavior"
			InitialValue=""
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
		#tag ViewProperty
			Name="DataSource"
			Visible=true
			Group="Database Binding"
			InitialValue=""
			Type="String"
			EditorType="DataSource"
		#tag EndViewProperty
		#tag ViewProperty
			Name="DataField"
			Visible=true
			Group="Database Binding"
			InitialValue=""
			Type="String"
			EditorType="DataField"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass

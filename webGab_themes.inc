<%
' this bit of script allows you to set a color
' theme and button style for webNav
' un-comment the theme or style you want to
' activate and comment the old one

dim strStyle, arColor

' Grayscale:
'arColor = Array("e0e0e0","d0d0d0","c0c0c0","b0b0b0","a0a0a0","909090","808080","303030","e0e0e0","ffffff","a0a0a0","e0e0e0","c0c0c0","a0a0a0","808080")

' Turqoise on white background:
'arColor = Array("dffcfc","ffffff","cfecec","bfdcdc","afcccc","9fbcbc","8facac","003355","ff7777","f0f0ff","99bbdd","e0e0e0","c0c0c0","a0a0a0","808080")

' Spring colors:
'arColor = Array("a9ff76","99ff66","89df56","79cf46","69bf36","59af26","499f16","ffffff","ffffdd","ffff44","000000","e0e0e0","c0c0c0","a0a0a0","808080")

' Brown colors:
'arColor = Array("ffdca9","ffcc99","efbc89","dfac79","cf9c69","bf8c59","af7c49","661100","00ccff","ffffff","cc9966","e0e0e0","c0c0c0","a0a0a0","808080")

' Purplish colors:
arColor = Array("efedff","dfddff","cfcdef","bfbddf","afadcf","9f9dbf","8f8daf","000000","440044","009966","880000","e0e0e0","c0c0c0","a0a0a0","808080")

' Orange colors:
' arColor = Array("ffb953","ff9933","df7913","bf5903","9f3902","7f1901","5f0900","330033","440044","ff6600","c0d0e0","e0e0e0","c0c0c0","a0a0a0","808080")
'                0        1        2        3        4        5        6        7        8        9        10       11       12       13       14

' here's a key to the colors as presently used:
' 0 =              [lightest shade]
' 1 = background   .
' 2 =              . 
' 3 =              .
' 4 = title text,  heading background
' 5 =              .
' 6 =              [darkest shade]
' 7 = message text [begin message colors]
' 8 = message link
' 9 = message active link
'10 = new message color
'[the following probably won't need to be changed]
'11 = light grey
'12 =     .
'13 =     .
'14 = dark grey

'style = "arrow"
strStyle = "box"
'style = "folder"

%>

<STYLE TYPE="text/css">
	A:hover { color:#<%=arColor(9)%>; }
	A:link { text-decoration: none; }
	A:active { text-decoration: none; }
	A:visited { text-decoration: none; }
</STYLE>
<pre>
Sub AddComment(Optional s)

    Set commentRange = ActiveCell
    
    commentRange.AddComment
    commentRange.Comment.Visible = False
    commentRange.Comment.Text Text:=s
End Sub
</pre>

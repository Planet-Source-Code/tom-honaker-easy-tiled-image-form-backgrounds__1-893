<div align="center">

## Easy tiled\-image form backgrounds


</div>

### Description



This code, which was inspired by a similar snippet of code by Ian Ippolito, permits tiling an image onto a form's background. This variant, though, resides in a module and is called by a form instead of residing within the form's code itself. This permits using the feature project-wide without redundant code all over the place.
 
### More Info
 


Three inputs are required in the actual sub call:

frm:    The name of the form to tile an image onto

picholder: The name of a PictureBox control on that same form, which will be used to hold the image to tile. (See the explanation for required settings)

bkgdfile: The image file to load, with complete path.

The best place to put the call I find is in the form's Form_Paint() event.



To use this code, you'll need to create a PictureBox control on your target form and set its AutoRedraw, AutoSize, and ClipControls properties to TRUE, and its Visible property to FALSE. When you call the sub, you'll be passing the form's name and the name of this PictureBox to the sub.



Nothing returned.



Larger forms (read: INSANELY large forms) might take some time to refresh on slow systems. Also, very small images take noticeably longer to tile than larger ones, so aim for about 125x125 pixel sizes for the tilable images.

Finally, although this code can use GIFs, interlaced GIFs tend to produce a single-pixel horizontal banding effect. So convert 'em to non-interlaced, etc.

Also, this does NOT work on MDI parent forms, although it works great on MDI child forms.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Tom Honaker](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tom-honaker.md)
**Level**          |Unknown
**User Rating**    |4.3 (186 globes from 43 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tom-honaker-easy-tiled-image-form-backgrounds__1-893/archive/master.zip)





### Source Code

```

Sub TileBkgd(frm As Form, picholder As PictureBox, bkgdfile As String)
  If bkgdfile = "" Then Exit Sub
  Dim ScWidth%, ScHeight%, ScMode%, n%, o%
  ScMode% = frm.ScaleMode
  picholder.ScaleMode = 3
  frm.ScaleMode = 3
  picholder.Picture = LoadPicture(bkgdfile)
  picholder.ScaleMode = 3
  For n% = 0 To frm.Height Step picholder.ScaleHeight
    For o% = 0 To frm.Width Step picholder.ScaleWidth
      frm.PaintPicture picholder.Picture, o%, n%
    Next o%
  Next n%
  frm.ScaleMode = ScMode%
  picholder.Picture = LoadPicture()
End Sub
```


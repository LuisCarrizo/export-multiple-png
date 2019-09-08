# export-multiple-png

Coreldraw (vba) macro to export a page to  multiple **ic-launcher.png** files according to standard size (hdpi, mdpi, xhdpi, xxhdpi and xxxhdpi plus 512x512 px).   

This macro export a page from a .cdr file, generating the following folders on the same folder of the .cdr file:
* mdpi    (48x48 px)
* hdpi    (72x72 px)
* xhdpi   (96x96 px)
* xxhdpi  (144x144 px)
* xxxhdpi (196x196 px)
* 512     (512x512 px)

The page to export must be named as **square**.

Additionaly, will export a page named **rounded** and generate **ic-launcher-round.png**.

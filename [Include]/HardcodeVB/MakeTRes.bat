@echo off
Rem Make US
rc /v /d US /r /fo tresus.res tres.rc
Rem Make Swinish
rc /v /d SW /r /fo tressw.res tres.rc
If "%1"=="" Goto Done
Rem Make version specific if country (US or SW) given on command line
If exist tres%1.res Copy tres%1.res tres.res
:Done

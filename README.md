# QT
There are some projects using by the QT5 platform.


This function of code is saving and creating excel. 
Attention, you should add some codes to choose the file path you want to save. 
For example, you can use this code
{savePath = QFileDialog::getSaveFileName(this, tr("Save orbit"), ".", tr("Microsoft Office 2010 (*.xlsx)")); } 
to create a public variable savePath to save your excel.

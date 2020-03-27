Development notes
=================

The exe relies on finding the dictionaries.  Rather than bundle these in the code:

1.  create a directory D:\Proyecto Libros\Diccionarios 
2.  unzip glare.rar in there, to create sub-directory glare with all the page images in
3.  do the same with liddell-scott

The code looks first for a folder Diccionarios under the exe dir; if it doesn't find it,
it looks for the directory above.

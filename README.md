# py-compare
Python Machine Learning database merger 


![image](https://user-images.githubusercontent.com/19597283/52545532-2b7d2480-2d86-11e9-9dfd-5a85d04e79a6.png)
 
 
 Machine Learning script turned into stand-alone executable application using [cx_freeze](https://anthony-tuininga.github.io/cx_Freeze/) and [tkinter](https://docs.python.org/2/library/tkinter.html).
 
### About


The application has the following features:

* It can be run from any computer in the organization without the need to install python.
* Using [multiprocessing](https://docs.python.org/2/library/multiprocessing.html), the application runs the gui and the calculations on different processors, allowing to trace the progress in real time using a progress bar
* The program uses an Excel template to export the results to facilitate the post-merge process.
* The program uses a machine learning algorithm with string distance calculations to decide if a pairing correspond to the same individual 



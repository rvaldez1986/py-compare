# py-compare
Python Machine Learning database merger 


![image](https://user-images.githubusercontent.com/19597283/52545532-2b7d2480-2d86-11e9-9dfd-5a85d04e79a6.png)
 
 
 Machine Learning script turned into stand-alone executable application using [cx_freeze](https://anthony-tuininga.github.io/cx_Freeze/) and [tkinter](https://docs.python.org/2/library/tkinter.html).
 
### About

An application that can be run from any computer in the organization without the need to install python. 

For the 

explores multi processing capabilities and object oriented programming. Multi processing is also used for the gui loading bar working in parallel with the computation. 

Finally it also uses an Excel template so it can be incorporated with prior and posterior analysis in this software.

Although the program uses a simple string distance algorithm for comparing among entries, the speed and user friendliness are optimized.

The application has the following features:

* It can be run from any computer in the organization without the need to install python.
* Using [multiprocessing](https://docs.python.org/2/library/multiprocessing.html), the application runs the gui and the calculations on different processors, allowing to trace the progress in real time using a progress bar
* send_generic_message(recipient_id, elements)
* send_button_message(recipient_id, text, buttons)
* send_attachment(recipient_id, attachment_type, attachment_path)
* send_attachment_url(recipient_id, attachment_type, attachment_url)
* send_image(recipient_id, image_path)
* send_image_url(recipient_id, image_url)
* send_audio(recipient_id, audio_path)
* send_audio_url(recipient_id, audio_url)
* send_video(recipient_id, video_path)



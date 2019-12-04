# py-compare
Python Machine Learning database merger 




![image](https://user-images.githubusercontent.com/19597283/52545532-2b7d2480-2d86-11e9-9dfd-5a85d04e79a6.png)
 
 
 Python Wrapper for [Facebook Messenger Platform](https://developers.facebook.com/docs/messenger-platform).
 
### About

project using tkinter, cx_freeze and pandas. 

The objective of the application is to create a stand-alone application that can be run from any computer in the organization. 
The application uses tkinter as the GUI, cx_freeze to create the application as stand-alone, explores multi processing capabilities and object oriented programming. Multi processing is also used for the gui loading bar working in parallel with the computation. 

Finally it also uses an Excel template so it can be incorporated with prior and posterior analysis in this software.

Although the program uses a simple string distance algorithm for comparing among entries, the speed and user friendliness are optimized.

This wrapper has the following functions:

* send_text_message(recipient_id, message)
* send_message(recipient_id, message)
* send_generic_message(recipient_id, elements)
* send_button_message(recipient_id, text, buttons)
* send_attachment(recipient_id, attachment_type, attachment_path)
* send_attachment_url(recipient_id, attachment_type, attachment_url)
* send_image(recipient_id, image_path)
* send_image_url(recipient_id, image_url)
* send_audio(recipient_id, audio_path)
* send_audio_url(recipient_id, audio_url)
* send_video(recipient_id, video_path)



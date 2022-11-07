import time
import win32com.client

### Shell Connection
shell = win32com.client.Dispatch("WScript.Shell")
shell.Run("notepad")
time.sleep(1)
shell.AppActivate("Notepad")

### Message
msg = """Science, engineering and technology
The distinction between science, engineering and technology is not always clear. 
Science is the reasoned investigation or study of phenomena, aimed at discovering enduring principles
among elements of the phenomenal world by employing formal techniques such as the scientific method. 
Technologies are not usually exclusively products of science, because they have to satisfy requirements 
such as utility, usability and safety.
Engineering is the goal-oriented process of designing and making tools and systems to exploit natural 
phenomena for practical human means, often using results and techniques from science. 
The development of technology may draw upon many fields of knowledge, including scientific, engineering, 
mathematical, linguistic, and historical knowledge, to achieve some practical result.

Technology is often a consequence of science and engineering — although technology as a human activity
precedes the two fields. For example, science might study the flow of electrons in electrical conductors, 
by using already-existing tools and knowledge. This new-found knowledge may then be used by engineers to 
create new tools and machines, such as semiconductors, computers, and other forms of advanced technology. 
In this sense, scientists and engineers may both be considered technologists; the three fields are often considered 
as one for the purposes of research and reference."""

###For Loop for sending characters in sequence with a delay
delay=0.04
for i in msg:
    time.sleep(delay)
    shell.SendKeys(i, 0)
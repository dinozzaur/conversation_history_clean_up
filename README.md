# conversation_history_clean_up
Outlook Conversation History folder duplicates clean up script

There is a known behavior with on-prem Exchange 2016 and Outlook 2013/2016 when Skype for Business client save multiple versions of the same conversation into Conversation History folder where all next versions of the conversation contain all the messages which were already saved previously plus new messages. By doing this it consumes a lot of space and generates a lot of useless items. Outlook built-in clean up tools are not abe to handle this. This script might be not perfect from coding point of view, hoewever it serves the purpose and I was looking for solution of the issue several years and was not able to find one. 
 
This script finds all items in conversation history folder which relate to single conversation where only last saved version of such conversations
contains all the messages and all previous saved versions of the same conversation obviously contain only message up to the time when it was saved and so they
are not needed and just consume storage space.

Script was tested with on-prem Exchange 2016 and Outlook 2016 client running Windows 10 LTSC build. It is not guranteed that it will run properly in another environment. 

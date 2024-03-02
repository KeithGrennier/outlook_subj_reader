Most of the work is done by this github "https://github.com/Hridai/Automating_Outlook/blob/master/ol_script.py" and is definetely worth the read.

I wanted to iterate email for headers. Such as '!!' 'emergency'.

While using this beware of not being to use outlook. Outlook updates running causing program to not run. Closing outlook crashes the program.

If you'd like to modify the search paramaters found in word_bank.json, just edit any of the words in the "words" section. If you'd like to change the folder, just modify it from "Inbox" to another folder you have such as "Spam"

```
{
    "words" : [
        "word1",
        "word2"
    },
    "folder":"Spam"
}
```

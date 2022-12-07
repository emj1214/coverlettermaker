# Author: Emma Johnson
# Date Last Updated: December 6, 2022
# URL: https://github.com/emj1214/coverlettermaker

myfirstname = "FIRST"
mylastname = "LAST"
myprefix = "Dr."
whodesc = "a third-year Dartmouth student"
myphone = "+X.XXX.XXX.XXXX"
myemail = "email@example.com"
myaddress = "STREET \nCITY, STATE ZIP"
mypronouns = "they/them/theirs"

# adjust and change these options as needed – be sure to replace the Xs before using
valopts = ["X skills", "ability to X", "unique set of X skills", "X", "X", "X", "strengths as defined by the Clifton StrengthsFinder Assessment"]


strengthssent = "" # use this string to list your Clifton Strengths and define how you use them

# feel free to change or remove this – it's just my go-to
learningsentp1 = "As a student, tutor, and teacher myself, learning is very important to me, and the kind of environment described is one that would foster my knowledge of "
learningsentp2 = " which in turn will lead to X, X, and X that would make working for " # replaces Xs with what learning does for you
learningsentp3 = " feel like anything but work." # feel free to change this – it's just my go-to

mylinks = ["URL1" ,"URL2", "URL3", "URL4"]
achieveopts = [
    ["Unlinked Text Portion 1", "Hyperlinked text", "Unlinked Text Portion 2"],
    ["Unlinked Text Portion 1", "Hyperlinked text", "Unlinked Text Portion 2"],
    ["Unlinked Text Portion 1", "Hyperlinked text", "Unlinked Text Portion 2"], 
    ["Unlinked Text Portion 1", "Hyperlinked text", "Unlinked Text Portion 2"]
]
achieveoptslist = []
for i in achieveopts :
    i = "".join(i)
    i = i[:99]
    achieveoptslist.append(i)
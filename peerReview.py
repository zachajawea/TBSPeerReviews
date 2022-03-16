# TODO: Add a roster class
# TODO: Add a means to ouput stats excel sheet for each squad and whole platoon

"""This program reads in a roster, an email list, and an excel sheet of
    unsorted peer review comments.

   It builds a list of students with their corresponding emails, sorts the list
   of comments, assigns the comments to the correct student, and then emails
   the student their own list of comments.

   Program also tells a student what their highest, lowest, and average
   ranking was from their peers."""
# Import the email modules
import csv
import smtplib
from email.mime.text import MIMEText
import sys
import math


class peerReview():
    """The main class of this program."""
    class student():
        """Sub class to hold information about a student. roster is a list of
           students."""

        def setEmail(self, email):
            """set the email of a student"""
            self.email = email

        def getEmail(self):
            """get the email of a student."""
            return self.email

        def getLastName(self):
            """return the student last name"""
            return self.lastName

        def getFirstName(self):
            """return the student first name"""
            return self.firstName

        def getComments(self):
            """Return the list of comments for a student"""
            return self.comments

        def getNumComments(self):
            """Returns the # of comments a student has"""
            return len(self.comments)

        def addComment(self, comment):
            """Appends a formatted comment to the end of the comments list"""
            self.comments.append(str(len(self.comments)+1)+") "+str(comment))

        def addRank(self, rank):
            """Add the rank a peer ranked this student"""
            self.allRanks.append(rank)

        def getHighestRank(self):
            """Returns the highest rank a student was ranked by peers"""
            return self.highestRank

        def getLowestRank(self):
            """Returns the lowest rank a student was ranked by peers"""
            return self.lowestRank

        def getAvgRank(self):
            """Returns the average rank a student was ranked by all peers"""
            return self.avgRank

        def calcAvgRank(self):
            """Calculates the average rank for a student given by all peers"""
            if len(self.allRanks) is 0:
                print("No ranks entered.")
                return False
            return sum(self.allRanks)/len(self.allRanks)

        def setHighestRank(self):
            """Searches self.allRanks for highest ranking,
                or the lowest # Value"""
            highest = 99
            for rank in self.allRanks:
                if rank < highest:
                    self.highestRank = rank

        def setLowestRank(self):
            """Searches self.allRanks for the lowest ranking,
                or the highest # value"""
            lowest = 0
            for rank in self.allRanks:
                if rank > lowest:
                    self.lowestRank = lowest

        def setEmailBody(self):
            """Builds the email body to be sent to a student. Includes all
            ranking metrics as well as all comments."""
            body = "Hello "
            body += self.lastName + ", " + self.firstName + ":\n\n"
            body += "Highest Rank: " + str(self.highestRank) + "\n"
            body += "Lowest Rank: " + str(self.lowestRank) + "\n"
            body += "Average Rank: " + str(self.avgRank) + "\n\n"
            body += "Listed below are your Peer Ranking Comments:\n"
            for comment in self.comments:
                body += comment + "\n"

            body += "If you have any questions or discrepancies, reply to \n"
            body += "this email from zachcammer@gmail.com"
            self.emailBody = body

        def getEmailBody(self):
            """Returns body of email prepared by setEmailBody"""
            return self.emailBody

        def __init__(self, name):
            self.email = ""
            self.lastName, self.firstName = name.split(',')
            self.lastName = self.lastName.strip()
            self.firstName = self.firstName.strip()
            self.comments = []
            self.allRanks = []
            self.highestRank = 0
            self.lowestRank = 0
            self.avgRank = 0
            self.sqd = 0
            self.emailBody = "Not Set"

    class roster():
        """Sub class.  List of Students and associated methods for that list"""
        def __init__(self):
            """creates the roster by filling a list called students, then
             alphabetically sorting that list."""
            self.students = []
            self.fill()
            self.sortAlpha()

        def report(self, size):
            """Method to create excel sheets of analytics for Capt."""
            if size == "platoon":
                excelFile = self.makeFile("5th Platoon")
            elif size == "squad":
                for student in self.roster.students.sqd == "1":
                    excelFile = self.makeFile("1st Squad")
                for student in self.roster.students.sqd == "2":
                    excelFile = self.makeFile("2nd Squad")
                for student in self.roster.students.sqd == "3":
                    excelFile = self.makeFile("3rd Squad")

        def makeFile(self, nameOfFile):
            """makes the excel file"""
            pass

        def sortAlpha(self):
            """Sorts list of students alphabetically."""
            pass

        def sortRank(self):
            """Sorts students by their avg rank from 1 - N, where N is the
             greatest rank."""
            pass

        def namesAndEmails(self):
            """Prints the name and email of all students"""
            for student in self.students:
                print(student.lastName + '\t' + student.email)
            return True

    def __init__(self):
        self.T = True
        self.V = False
        self.gmail_user = input("Enter Username:")
        self.gmail_pwd = input("Enter Password:")

    def login(self):
        """login to gmail using specified uname and pwd"""

        if self.V:
            print("Logging in to: "+self.gmail_user)
        if self.T is not True:
            server_ssl = smtplib.SMTP_SSL("smtp.gmail.com", 465)
            server_ssl.ehlo()  # optional, called by login()
            if self.V:
                print("Login Successful.")
                print("Logged into "+self.gmail_user+".")
            return server_ssl.login(self.gmail_user, self.gmail_pwd)
        if self.T is True:
            if self.V:
                print("Running in \"Troubleshoot\" mode, notionally logged in")
            return ""

    def sendEmail(self, server_ssl, recipient):
        """function to send an email.  Only needs the contents and who will be
           the receiving email"""
        try:
            recipient.setEmailBody()
            if self.T:
                print(recipient.getEmailBody())
            if self.V:
                print("Sending Email to: " + recipient.getLastName)
                print("Email body is: " + recipient.getEmailBody)
            FROM = self.gmail_user
            recipNm = recipient.getLastName()
            TO = recipNm if isinstance(recipNm, list) is list else [recipNm]
            SUBJECT = "Peer Ranking Comments"
            TEXT = recipient.getEmailBody()

            # Prepare actual message
            message = """From: %s\nTo: %s\nSubject: %s\n\n%s
            """ % (FROM, ", ".join(TO), SUBJECT, TEXT)

            if self.T is False:
                # ssl server doesn't support or need tls
                server_ssl.sendmail(FROM, TO, message.encode('utf-8'))
                # server_ssl.quit()
                print('Successfully sent an email to ' + recipNm)
            if self.T is True:
                print("Notionally sent an email to " + recipNm)
        except Exception as e:
            print("Emails not sent. Error to Follow.")
            print(e)
            return False
        return True

    def mailingList(self, roster):
        """Takes roster from main and adds an email address to each student.
           Emails can be from text file or user entered as comma separated
           list."""
        print("\nBuilding Mailing List")
        # Get List of Emails to send message to
        if self.V:
            print("User will be prompted to either enter a (.txt) file to" +
                  "\nhave the program read the student emails from.")
        try:
            fname = input("File to import from:" + '\n' + "> ")
            emails = []
            with open(fname) as fin:
                for line in fin:
                    emails.append(line)
            for i in range(len(roster)):
                email = emails[i].strip()
                roster[i].setEmail(email)
                if self.V:
                    print("Student, " + roster[i].getLastName() +
                          ", assigned email: " + email)
        except IOError as e:
            print(str(e))
            return False

    def sortPeerEvals(self, roster, unsortedComments):
        """Adds unsorted Comments to the Student that they belong to in
           roster. Also adds a rank then averages all ranks given to student
           at the end."""
        for student in roster:
            name = student.getLastName().upper()
            print("Looking for "+name+" comments.")
            for comment in unsortedComments:
                if comment[0] == name:
                    if self.T:
                        print("Found one.")
                    student.addComment(comment)
                    student.addRank(comment[2])
            student.calcAvgRank()


    def setProgramOptions(self):
        """User can set program to run in troubleshoot which provides more
           information for fixing operation, or verbose, which provides
           more information for correct usage of program."""
        print("Comma seperated run options:")
        print("(Verbose, Troubleshoot, N/A)")
        options = input("> ")
        options = options.split(',')

        for option in options:
            option = option.strip()

        for option in options:
            option = option.upper()
            if option == "NA" or option == "N/A":
                print("NO OPTIONS SET.")
                return
            if option == "VERBOSE":
                self.V = True
                print("Running in verbose mode.")
            if option == "TROUBLESHOOT":
                self.T = True
                print("Running in troubleshoot mode.")

    def displayWarningPage(self):
        """Query user to see if program has necessary files to run."""
        print("//////////////////////////////////////////////////////////////")
        print("///////////////////////PLEASE READ////////////////////////////")
        print("// User should ensure that the following files are in the   //")
        print("// same folder as this program:                             //")
        print("// - Roster file (.txt)                                     //")
        print("// - emailing list in same order as roster file (.txt)      //")
        print("// - the excel spreadsheet of comments (.csv)               //")
        print("// Does the user have the above files in the SAME FOLDER AS //")
        print("// this program?                                            //")
        print("//////////////////////////////////////////////////////////////")
        ans = input("Answer:\n1) Yes\n2) No\n> ")
        if ans == "1" or ans == "yes" or ans == "Yes" or ans == "YES":
            return True
        print("Please set up the above files and run this program again.")
        return False

    def emailPeerEvals(self, roster):
        if self.T is False:
            server_ssl = peerReview.login(self)
        else:
            server_ssl = None
        for student in roster:
            emailSent = peerReview.sendEmail(self, server_ssl, student)
            if emailSent is not True:
                return False
        if self.T is False:
            server_ssl.close()
        return True

    def rosterBuilder(self):
        """Build the list of students from a text file"""
        if self.V:
            print("\nBuilding the student roster for the program.")
            print("You will need to provide a (.txt) roster file. Roster " +
                  "file\nshould have (1) name (Last Name,First Name) per line")
        fname = input("What is the roster filename?\n> ")
        programRoster = []
        try:
            with open(fname, "r") as rosterFile:
                for name in rosterFile:
                    newStudent = peerReview.student(name)
                    programRoster.append(newStudent)
                    if self.T:
                        print("Made new student, "+newStudent.getLastName() +
                              ", and added them to roster.")
        except IOError as e:
            if self.T and not self.V:
                print("File Read Error: "+str(e))
            if self.V:
                print("File Read Error. Exiting Program.")
            return False
        return programRoster

    # dont touch this until everything else works.
    def rosterBuilderV2(self):
        """use one excel roster sheet for email list, roster, and squads"""
        pass

    def nameCommentList(self):
        """reads from an excel spreadsheet of unorganized peer review comments,
           and builds a list of name:comment:rank triples.  Name is of person
           comment is about.  Returns false for file read errors."""
        listOfComments = []
        try:
            print("What is the name of the peer review comments Excel file?")
            fname = input("> ")
            with open(fname, newline='', encoding='utf-8') as csvfile:
                spamreader = csv.reader(csvfile)
                for row in spamreader:
                    for i in range(0, len(row), 2):
                        if self.V:
                            print(str(row[i])+'\t'+str(row[i+1]))
                        listOfComments.append((row[i].strip().upper(),
                                               row[i+1].strip().upper(),
                                               math.ceil(i/2)))
        except IOError as e:
            redo = input(str(e)+'\n'+"Try different file?\n1) Yes\n2) No\n> ")
            redo = redo.upper()
            if redo is "1" or redo is "YES":
                peerReview.nameCommentList(self)
            else:
                return False
        return listOfComments

    def check(self, roster):
        """Make sure every student has an email. Prompt the User to double check
        the emails assigned to each Student."""
        for student in roster:
            print("Student, " + student.getLastName() +
                  ", assigned email: " + student.getEmail())
        print("Are all of the above Name : Email combinations correct?")
        print("1) Yes \n2) No")
        ans = input("> ")
        if ans == "no":
            return False
        return True


if __name__ == "__main__":
    pr = peerReview()
    # Choose to run Verbose or Troubleshoot
    pr.setProgramOptions()
    # Warn user of necessary documents needed for program to run successfully
    if pr.displayWarningPage() is False:
        sys.exit()
    # Build the list of students
    roster = pr.rosterBuilder()
    # Add emails to students
    pr.mailingList(roster)
    # Seperate Excel Sheet into dictionary of name : comment
    unsortedComments = pr.nameCommentList()
    # Build List of comments for each student
    pr.sortPeerEvals(roster, unsortedComments)
    # Check # of students in roster against emails and comments per student
    if pr.check(roster) is False:
        sys.exit()
    # Email each student their peer ranking comments / high, low, & avg rank
    pr.emailPeerEvals(roster)
    input("press any button to exit")
    sys.exit()

namespace Renamer4
{
	internal static class Program
	{
		[STAThread]
		static void Main()
		{
			Console.WriteLine("Opening Dialog");

			//Open Folder Dialog
			FolderBrowserDialog folderDialog = new FolderBrowserDialog();
			folderDialog.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory;
			if (folderDialog.ShowDialog() == DialogResult.OK)
			{
				string directory = folderDialog.SelectedPath;
				Console.WriteLine("Selected Path: " + directory);

				string[] fileArray = Directory.GetFileSystemEntries(directory);
				var shellAppType = Type.GetTypeFromProgID("Shell.Application");
				if(shellAppType != null)
				{
					dynamic? shell = Activator.CreateInstance(shellAppType);
					if(shell != null)
					{
						for (int i = 0; i < fileArray.Length; i++)
						{
							string filePath = fileArray[i];
							string fileName = Path.GetFileName(filePath);
							FileAttributes attr = File.GetAttributes(filePath);	//Necessary Evil
							if (!attr.HasFlag(FileAttributes.Directory))		//Not a directory, go ahead and change the fileName
							{
								//Console.WriteLine("Converting \"" + fileName+"\"");

								Shell32.Folder objFolder = shell.NameSpace(Path.GetDirectoryName(filePath));				//Necessary Evil
								Shell32.FolderItem folderItem = objFolder.ParseName(Path.GetFileName(filePath));	//Necessary Evil
								String ogDateString = objFolder.GetDetailsOf(folderItem, 208); //208 = Media Created
								if(ogDateString.Length == 0)
									ogDateString = objFolder.GetDetailsOf(folderItem, 12); //12 = Date Taken
								if (ogDateString.Length > 0) //Go ahead if it has a "Media Created" date
								{
									String dateTimeString = parseString(ogDateString);
									String newFileName = dateTimeString + "_" + fileName;
									Console.WriteLine("renaming \"" + fileName + "\"\t to \t\"" + newFileName + "\"");
									renameFile(directory, fileName, newFileName);
								}
							}
							else
							{
								Console.WriteLine(fileName + " is not a file");
							}
						}
					}
					else
					{
						Console.WriteLine("ERROR! shell=null");
					}
				}
				else
				{
					Console.WriteLine("ERROR! shellAppType=null");
				}
			}
		}

		static void renameFile(string directory, string oldFilename, string newFilename)
		{
			string oldFullPath = directory + "\\" + oldFilename;
			string newFilePath = directory + "\\" + newFilename;
			File.Move(oldFullPath, newFilePath);
		}

		static string parseString(string ogDateString)
		{
			char[] ogDateChars = ogDateString.ToCharArray();
			int monthStart = 1;
			int monthLength = 2;
			String monthString = "";
			if (ogDateChars[monthStart + monthLength - 1] == '/')
			{
				monthLength = 1;
				monthString += "0";
			}
			monthString += ogDateString.Substring(monthStart, monthLength);

			int dayStart = monthStart + monthLength + 2;
			int dayLength = 2;
			String dayString = "";
			if (ogDateChars[dayStart + dayLength - 1] == '/')
			{
				dayLength = 1;
				dayString += "0";
			}
			dayString += ogDateString.Substring(dayStart, dayLength);

			int yearStart = dayStart + dayLength + 2;
			int yearLength = 4;
			String yearString = ogDateString.Substring(yearStart, yearLength);

			int hourStart = yearStart + yearLength + 3;
			int hourLength = 2;
			if (ogDateChars[hourStart + hourLength - 1] == ':')
				hourLength = 1;
			String hourString = ogDateString.Substring(hourStart, hourLength);
			int hourInt = int.Parse(hourString);

			int minuteStart = hourStart + hourLength + 1;
			int minuteLength = 2;
			String minuteString = ogDateString.Substring(minuteStart, minuteLength);

			int amPmStart = minuteStart + minuteLength + 1;
			if (ogDateChars[amPmStart] == 'A' && hourInt == 12)
				hourInt = 0;
			else if (ogDateChars[amPmStart] == 'P' && hourInt != 12)
				hourInt += 12;
			hourString = "";
			if (hourInt < 10)
				hourString += "0";
			hourString += hourInt.ToString();

			string finalString = yearString + monthString + dayString + hourString + minuteString;
			//Console.WriteLine(yearString + " " + monthString + " " + dayString + " " + hourString + " " + minuteString + " | " + ogDateString);
			return finalString;
		}
	}
}

/*
  0 Name
  1 Size
  2 Item type
  3 Date modified
  4 Date created
  5 Date accessed
  6 Attributes
  7 Offline status
  8 Availability
  9 Perceived type
 10 Owner
 11 Kind
 12 Date taken
 13 Contributing artists
 14 Album
 15 Year
 16 Genre
 17 Conductors
 18 Tags
 19 Rating
 20 Authors
 21 Title
 22 Subject
 23 Categories
 24 Comments
 25 Copyright
 26 #
 27 Length
 28 Bit rate
 29 Protected
 30 Camera model
 31 Dimensions
 32 Camera maker
 33 Company
 34 File description
 35 Masters keywords
 36 Masters keywords
 42 Program name
 43 Duration
 44 Is online
 45 Is recurring
 46 Location
 47 Optional attendee addresses
 48 Optional attendees
 49 Organizer address
 50 Organizer name
 51 Reminder time
 52 Required attendee addresses
 53 Required attendees
 54 Resources
 55 Meeting status
 56 Free/busy status
 57 Total size
 58 Account name
 60 Task status
 61 Computer
 62 Anniversary
 63 Assistant's name
 64 Assistant's phone
 65 Birthday
 66 Business address
 67 Business city
 68 Business country/region
 69 Business P.O. box
 70 Business postal code
 71 Business state or province
 72 Business street
 73 Business fax
 74 Business home page
 75 Business phone
 76 Callback number
 77 Car phone
 78 Children
 79 Company main phone
 80 Department
 81 E-mail address
 82 E-mail2
 83 E-mail3
 84 E-mail list
 85 E-mail display name
 86 File as
 87 First name
 88 Full name
 89 Gender
 90 Given name
 91 Hobbies
 92 Home address
 93 Home city
 94 Home country/region
 95 Home P.O. box
 96 Home postal code
 97 Home state or province
 98 Home street
 99 Home fax
100 Home phone
101 IM addresses
102 Initials
103 Job title
104 Label
105 Last name
106 Mailing address
107 Middle name
108 Cell phone
109 Nickname
110 Office location
111 Other address
112 Other city
113 Other country/region
114 Other P.O. box
115 Other postal code
116 Other state or province
117 Other street
118 Pager
119 Personal title
120 City
121 Country/region
122 P.O. box
123 Postal code
124 State or province
125 Street
126 Primary e-mail
127 Primary phone
128 Profession
129 Spouse/Partner
130 Suffix
131 TTY/TTD phone
132 Telex
133 Webpage
134 Content status
135 Content type
136 Date acquired
137 Date archived
138 Date completed
139 Device category
140 Connected
141 Discovery method
142 Friendly name
143 Local computer
144 Manufacturer
145 Model
146 Paired
147 Classification
148 Status
149 Status
150 Client ID
151 Contributors
152 Content created
153 Last printed
154 Date last saved
155 Division
156 Document ID
157 Pages
158 Slides
159 Total editing time
160 Word count
161 Due date
162 End date
163 File count
164 File extension
165 Filename
166 File version
167 Flag color
168 Flag status
169 Space free
172 Group
173 Sharing type
174 Bit depth
175 Horizontal resolution
176 Width
177 Vertical resolution
178 Height
179 Importance
180 Is attachment
181 Is deleted
182 Encryption status
183 Has flag
184 Is completed
185 Incomplete
186 Read status
187 Shared
188 Creators
189 Date
190 Folder name
191 Folder path
192 Folder
193 Participants
194 Path
195 By location
196 Type
197 Contact names
198 Entry type
199 Language
200 Date visited
201 Description
202 Link status
203 Link target
204 URL
208 Media created
209 Date released
210 Encoded by
211 Episode number
212 Producers
213 Publisher
214 Season number
215 Subtitle
216 User web URL
217 Writers
219 Attachments
220 Bcc addresses
221 Bcc
222 Cc addresses
223 Cc
224 Conversation ID
225 Date received
226 Date sent
	*/
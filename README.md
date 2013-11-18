Password Letters via iTextSharp and PowerShell
==============================================

At the ZHAW accounts are created through an automated process. The corresponding encrypted one-time-password and some metadata are stored in an SQL database. Later on HR staffer can create corresponding PDF password letters on request and provide them to end users.

The ZHAW uses the Adaxes Active Directory management and automation tool (www.adaxes.net) to provide certain AD management capabilities via a web interface to authorized users. This tool allows creation of different AD management web interfaces for specific user groups.

The task was, to extend the default Adaxes functionality with a non-standard functionality, allowing authorized users (HR) to request password letters in order for them to provide the password letters to end users.

The needed functionality was implemented through a PowerShell module using the AGPL version of the iTextSharp library (www.itextpdf.com). As requested by the license the implemented PowerShell module is released under the AGPL license. 

While implementing the needed functionality in PowerShell it became evident, that there is no good example or description, how to use iTextSharp in PowerShell. Most of the found examples focus on using the iTextSharp library in a C# application or ASP.Net web. We hope that the provided library and the example code can be used by others, when trying to create PDFs in PowerShell using iTextSharp.

This example includes the GeneratePWLetter and a Logging PowerShell module, an example script showing the usage of the module and a script for generating an example SQL database including some test data. Documentation can be found under \Documentation\ 131025-PW-Letter-AGPL.pdf.

This PowerShell module and the example are provided as it is and is free code. The ZHAW does not take any responsibility or provide any support.

Philipp Hungerbühler
(05.11.2013)


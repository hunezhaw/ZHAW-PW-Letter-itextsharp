truncate table tPWHistory
truncate table tPWAllADAccounts
truncate table tPWAllStudInClassesLocal

INSERT tPWHistory (TimeStamp,UniqueID,KurzZeichen,Departement,DepDescription,PasswordEncrypted) 
VALUES (GETDATE(), 123456, 'test1', 'X', 'DepX', EncryptByPassPhrase('YOUR-PW-DECRYPTION-KEY','test'))

INSERT tPWHistory (TimeStamp,UniqueID,KurzZeichen,Departement,DepDescription,PasswordEncrypted) 
VALUES (GETDATE(), 789123, 'test2', 'X', 'DepX', EncryptByPassPhrase('YOUR-PW-DECRYPTION-KEY','test'))

INSERT INTO tPWAllADAccounts (WhenCreated,PwdLastSet,LastLogon,UniqueID,KurzZeichen,PersKat,Nachname,Vorname,Geschlecht,EMail1,EMail2,Departement,DepDescription)
VALUES (GETDATE(), GETDATE(), null, 123456, 'test1', '000000', 'Test1LastName', 'Test1FirstName', 'M', 'test1@domain.com', null, 'X', 'DepX')

INSERT INTO tPWAllADAccounts (WhenCreated,PwdLastSet,LastLogon,UniqueID,KurzZeichen,PersKat,Nachname,Vorname,Geschlecht,EMail1,EMail2,Departement,DepDescription)
VALUES (GETDATE(), GETDATE(), null, 789123, 'test2', '#weiter#', 'Test2LastName', 'Test2FirstName', 'M', 'test2@domain.com', null, 'X', 'DepX')

INSERT INTO tPWAllStudInClassesLocal (UniqueID, KurzZeichen, Departement, DepDescription, AnlassEvento)
VALUES (123456, 'test1', 'X', 'DepX-EVENT-01', 'Event-01')
INSERT INTO tPWAllStudInClassesLocal (UniqueID, KurzZeichen, Departement, DepDescription, AnlassEvento)
VALUES (123456, 'test1', 'X', 'DepX-EVENT-02', 'Event-02')

INSERT INTO tPWAllStudInClassesLocal (UniqueID, KurzZeichen, Departement, DepDescription, AnlassEvento)
VALUES (789123, 'test2', 'X', 'DepX-EVENT-01', 'Event-01')
INSERT INTO tPWAllStudInClassesLocal (UniqueID, KurzZeichen, Departement, DepDescription, AnlassEvento)
VALUES (789123, 'test2', 'X', 'DepX-EVENT-03', 'Event-03')
DROP DATABASE IF EXISTS bctest;             
                                            
CREATE DATABASE bctest;                     
USE bctest;                                 
DROP TABLE IF EXISTS jobCard;               
CREATE TABLE jobCard(                       
    ID INT(10) NOT NULL AUTO_INCREMENT PRIMARY KEY,
    ClientName VARCHAR(50) NOT NULL,
    PhoneAndEmail VARCHAR(100) NOT NULL,
    Pass VARCHAR(30),
    Parts VARCHAR(100),
    ItemsAndServiced VARCHAR(100),
    Address VARCHAR(100),
    Issue VARCHAR(100) NOT NULL,
    Done VARCHAR(100) NOT NULL,
    Misc VARCHAR(100),
    Invoice VARCHAR(20)
    );                                      
ALTER TABLE jobcard AUTO_INCREMENT=4000;
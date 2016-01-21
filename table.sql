CREATE TABLE `plugin_patching_data` (
  `ID` int(11) NOT NULL AUTO_INCREMENT,
  `ComputerID` int(11) NOT NULL,
  `KB_Number` int(10) NOT NULL,
  `KB_Title` varchar(256) NOT NULL,
  `KB_Category` varchar(32) NOT NULL,
  `KB_Description` varchar(2048) NOT NULL,
  `KB_GUID` varchar(36) NOT NULL,
  `Severity` varchar(16) NOT NULL,
  `Result` varchar(16) NOT NULL,
  `hResult` varchar(16) NOT NULL,
  `hResultDesc` varchar(256) NOT NULL,
  `DateInstalled` datetime NOT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=InnoDB AUTO_INCREMENT=992 DEFAULT CHARSET=utf8
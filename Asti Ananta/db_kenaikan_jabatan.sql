# Host: localhost  (Version 5.5.5-10.1.13-MariaDB)
# Date: 2020-06-30 01:38:39
# Generator: MySQL-Front 5.3  (Build 5.33)

/*!40101 SET NAMES latin1 */;

#
# Structure for table "tabel_login"
#

DROP TABLE IF EXISTS `tabel_login`;
CREATE TABLE `tabel_login` (
  `username` int(11) NOT NULL AUTO_INCREMENT,
  `pasword` varchar(255) DEFAULT NULL,
  `jabatan` varchar(255) DEFAULT NULL,
  PRIMARY KEY (`username`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 ROW_FORMAT=COMPACT;

#
# Data for table "tabel_login"
#


#
# Structure for table "tabel_nilaiakhir"
#

DROP TABLE IF EXISTS `tabel_nilaiakhir`;
CREATE TABLE `tabel_nilaiakhir` (
  `nik` int(11) NOT NULL AUTO_INCREMENT,
  `namakaryawan` varchar(255) DEFAULT NULL,
  `jabatan` varchar(255) DEFAULT NULL,
  `bagian` varchar(255) DEFAULT NULL,
  `alamat` varchar(255) DEFAULT NULL,
  `jeniskelamin` varchar(255) DEFAULT NULL,
  PRIMARY KEY (`nik`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 ROW_FORMAT=COMPACT;

#
# Data for table "tabel_nilaiakhir"
#


#
# Structure for table "tabel_nilaikriteria"
#

DROP TABLE IF EXISTS `tabel_nilaikriteria`;
CREATE TABLE `tabel_nilaikriteria` (
  `namakaryawan` varchar(255) NOT NULL DEFAULT '',
  `jabatan` varchar(255) DEFAULT NULL,
  `masukkankriteria` varchar(255) DEFAULT NULL,
  `masakerja` varchar(255) DEFAULT NULL,
  `penilaiankinerja` varchar(255) DEFAULT NULL,
  `perilaku` varchar(255) DEFAULT NULL,
  PRIMARY KEY (`namakaryawan`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 ROW_FORMAT=COMPACT;

#
# Data for table "tabel_nilaikriteria"
#


-- ----------------------------------------------------------
-- MDB Tools - A library for reading MS Access database files
-- Copyright (C) 2000-2011 Brian Bruns and others.
-- Files in libmdb are licensed under LGPL and the utilities under
-- the GPL, see COPYING.LIB and COPYING files respectively.
-- Check out http://mdbtools.sourceforge.net
-- ----------------------------------------------------------

-- That file uses encoding UTF-8

CREATE TABLE [仕掛]
 (
	[ﾛｯﾄ番号]			Text (255) NOT NULL, 
	[引当番号]			Text (255) NOT NULL, 
	[検査番号]			Text (255) NOT NULL, 
	[受注番号]			Text (255) NOT NULL, 
	[出荷日]			Text (255) NOT NULL, 
	[受注_納入先ｺｰﾄﾞ]			Text (255) NOT NULL, 
	[受注_取引先ｺｰﾄﾞ]			Text (255) NOT NULL, 
	[受注_納入先名称]			Text (255) NOT NULL, 
	[受注_取引先名称]			Text (255) NOT NULL, 
	[引当数量]			Text (255) NOT NULL, 
	[引当調整NO]			Text (255) NOT NULL, 
	[引当ﾌｧｲﾙ登録年月日時分秒]			Text (255) NOT NULL
);



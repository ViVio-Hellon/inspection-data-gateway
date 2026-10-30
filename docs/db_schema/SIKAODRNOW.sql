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
	[受注番号]			Text (255) NOT NULL, 
	[VC_表]			Text (255) NOT NULL, 
	[VC_裏]			Text (255) NOT NULL, 
	[EX_輸出区分]			Text (255) NOT NULL, 
	[梱包単位_梱包形状_輸出先国名]			Text (255) NOT NULL, 
	[取引先名称]			Text (255) NOT NULL, 
	[納入先名称]			Text (255) NOT NULL, 
	[送り先名称]			Text (255) NOT NULL, 
	[包装仕様NO]			Text (255) NOT NULL, 
	[納送用ｺﾒﾝﾄ]			Text (255) NOT NULL, 
	[工場用ｺﾒﾝﾄ]			Text (255) NOT NULL, 
	[受注材質]			Text (255) NOT NULL, 
	[受注調質]			Text (255) NOT NULL, 
	[受注板厚]			Text (255) NOT NULL, 
	[受注板幅]			Text (255) NOT NULL, 
	[受注板丈]			Text (255) NOT NULL, 
	[用途ｺｰﾄﾞ]			Text (255) NOT NULL, 
	[用途名]			Text (255) NOT NULL, 
	[梱包ｺｰﾄﾞ]			Text (255) NOT NULL, 
	[製品単重]			Text (255) NOT NULL, 
	[材質_比重]			Text (255) NOT NULL, 
	[合紙]			Text (255) NOT NULL, 
	[ｺｲﾙ外径_MAX]			Text (255) NOT NULL, 
	[ｺｲﾙ外径_目標]			Text (255) NOT NULL, 
	[ｺｲﾙ外径_MIN]			Text (255) NOT NULL, 
	[コイル内径_目標]			Text (255) NOT NULL, 
	[梱包単位_重量]			Text (255) NOT NULL, 
	[梱包単位_枚数]			Text (255) NOT NULL, 
	[梱包単位_形態制限_総高さ]			Text (255) NOT NULL, 
	[営業納期]			Text (255) NOT NULL
);



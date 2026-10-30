-- ----------------------------------------------------------
-- MDB Tools - A library for reading MS Access database files
-- Copyright (C) 2000-2011 Brian Bruns and others.
-- Files in libmdb are licensed under LGPL and the utilities under
-- the GPL, see COPYING.LIB and COPYING files respectively.
-- Check out http://mdbtools.sourceforge.net
-- ----------------------------------------------------------

-- That file uses encoding UTF-8

CREATE TABLE [設備名バーコード]
 (
	[NO]			Text (255) NOT NULL, 
	[名前]			Text (255) NOT NULL, 
	[よみ]			Text (255) NOT NULL, 
	[バーコード]			Text (255) NOT NULL, 
	[番号]			Text (255), 
	[更新日]			Text (255)
);

CREATE TABLE [名簿]
 (
	[番号]			Text (255) NOT NULL, 
	[NO]			Text (255) NOT NULL, 
	[名前]			Text (255) NOT NULL, 
	[よみ]			Text (255) NOT NULL, 
	[バーコード]			Text (255) NOT NULL, 
	[NO加工無]			Text (255) NOT NULL, 
	[形態]			Text (255) NOT NULL, 
	[更新日]			Text (255) NOT NULL, 
	[管理者権限]			Boolean NOT NULL
);

CREATE TABLE [測定機器ID]
 (
	[ライン]			Text (255) NOT NULL, 
	[機器名]			Text (255) NOT NULL, 
	[ID]			Text (255) NOT NULL, 
	[出力テキスト]			Text (255) NOT NULL, 
	[厚下]			Text (255) NOT NULL, 
	[厚上]			Text (255) NOT NULL, 
	[用途]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL, 
	[更新日]			Text (255) NOT NULL, 
	[番号]			Long Integer NOT NULL
);

CREATE TABLE [時間用]
 (
	[番号]			Text (255) NOT NULL, 
	[直]			Text (255) NOT NULL, 
	[開始]			Text (255) NOT NULL, 
	[終了]			Text (255) NOT NULL, 
	[更新日]			Text (255) NOT NULL
);

CREATE TABLE [校正ストレートエッヂ]
 (
	[番号]			Text (255) NOT NULL, 
	[ライン名]			Text (255) NOT NULL, 
	[長さ]			Text (255) NOT NULL, 
	[長さ下]			Text (255) NOT NULL, 
	[長さ上]			Text (255) NOT NULL, 
	[許容範囲]			Text (255) NOT NULL, 
	[更新日]			Text (255) NOT NULL
);

CREATE TABLE [作業停止時間内訳_1]
 (
	[管理番号]			Text (255) NOT NULL, 
	[内訳]			Text (255) NOT NULL, 
	[内訳番号]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL
);

CREATE TABLE [設備仕様]
 (
	[管理番号]			Text (255) NOT NULL, 
	[設備名]			Text (255) NOT NULL, 
	[製造材質]			Text (255) NOT NULL, 
	[オーダー厚MIN]			Text (255) NOT NULL, 
	[オーダー厚MAX]			Text (255) NOT NULL, 
	[オーダー幅MIN]			Text (255) NOT NULL, 
	[オーダー幅MAX]			Text (255) NOT NULL, 
	[オーダー丈MIN]			Text (255) NOT NULL, 
	[オーダー丈MAX]			Text (255) NOT NULL, 
	[単重MIN]			Text (255) NOT NULL, 
	[単重MAX]			Text (255) NOT NULL, 
	[作業可否]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL
);

CREATE TABLE [VC厚]
 (
	[番号]			Text (255) NOT NULL, 
	[品名ｺｰﾄﾞ]			Text (255) NOT NULL, 
	[厚-係数]			Text (255) NOT NULL, 
	[厚]			Text (255) NOT NULL, 
	[系統]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL, 
	[貼り付け係数]			Text (255) NOT NULL, 
	[更新日]			Text (255) NOT NULL
);

CREATE TABLE [PkgSpecData]
 (
	[管理番号]			Text (255) NOT NULL, 
	[包装仕様書№]			Text (255) NOT NULL, 
	[版数]			Text (255) NOT NULL, 
	[旧包装仕様書№]			Text (255) NOT NULL, 
	[納入仕様書№]			Text (255) NOT NULL, 
	[梱包形態]			Text (255) NOT NULL, 
	[製品仕様]			Text (255) NOT NULL, 
	[梱包指定別]			Text (255) NOT NULL, 
	[タイトル]			Text (255) NOT NULL, 
	[発行日]			Text (255) NOT NULL, 
	[用途名]			Text (255) NOT NULL, 
	[代表サイズＴ]			Text (255) NOT NULL, 
	[代表サイズＷ]			Text (255) NOT NULL, 
	[代表サイズＬ]			Text (255) NOT NULL, 
	[仕様種別]			Text (255) NOT NULL, 
	[コメント]			Text (255) NOT NULL
);

CREATE TABLE [作業停止時間内訳_2]
 (
	[管理番号]			Text (255) NOT NULL, 
	[内訳]			Text (255) NOT NULL, 
	[内訳番号]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL
);

CREATE TABLE [作業停止時間内訳_3]
 (
	[管理番号]			Text (255) NOT NULL, 
	[内訳]			Text (255) NOT NULL, 
	[内訳番号]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL
);

CREATE TABLE [用途名負荷係数算出]
 (
	[管理番号]			Text (255) NOT NULL, 
	[用途名]			Text (255) NOT NULL, 
	[用途コード]			Text (255) NOT NULL, 
	[ロット数]			Text (255) NOT NULL, 
	[平均作業時間]			Text (255) NOT NULL, 
	[合計作業時間]			Text (255) NOT NULL, 
	[合計段取り時間]			Text (255) NOT NULL, 
	[合計作業時間＋合計段取り時間]			Text (255) NOT NULL, 
	[負荷係数]			Text (255) NOT NULL, 
	[PENAを1とした場合の換算係数]			Text (255) NOT NULL, 
	[係数丸め]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL
);

CREATE TABLE [キャスク梱包板番号]
 (
	[LotNo]			Text (255), 
	[検査No]			Text (255), 
	[スタートNo]			Long Integer, 
	[エンドNo]			Long Integer, 
	[作業時間]			DateTime, 
	[管理番号]			Long Integer
);



-- ----------------------------------------------------------
-- MDB Tools - A library for reading MS Access database files
-- Copyright (C) 2000-2011 Brian Bruns and others.
-- Files in libmdb are licensed under LGPL and the utilities under
-- the GPL, see COPYING.LIB and COPYING files respectively.
-- Check out http://mdbtools.sourceforge.net
-- ----------------------------------------------------------

-- That file uses encoding UTF-8

CREATE TABLE [中歪4級]
 (
	[平坦度上限値参照元]			Text (255) NOT NULL, 
	[用途コード]			Text (255) NOT NULL, 
	[板厚下]			Text (255) NOT NULL, 
	[板厚上]			Text (255) NOT NULL, 
	[500]			Text (255) NOT NULL, 
	[4000]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL, 
	[管理番号]			Text (255)
);

CREATE TABLE [QA_022_01]
 (
	[VC]			Text (255) NOT NULL, 
	[1m当たり平坦度上限]			Text (255) NOT NULL, 
	[板厚下]			Text (255) NOT NULL, 
	[板厚上]			Text (255) NOT NULL, 
	[厚]			Text (255) NOT NULL, 
	[500]			Text (255) NOT NULL, 
	[750]			Text (255) NOT NULL, 
	[1005]			Text (255) NOT NULL, 
	[1255]			Text (255) NOT NULL, 
	[1530]			Text (255) NOT NULL, 
	[2005]			Text (255) NOT NULL, 
	[2505]			Text (255) NOT NULL, 
	[3055]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL
);

CREATE TABLE [~TMPCLP491581]
 (
	[欠陥名]			Text (255) NOT NULL, 
	[1]			Text (255) NOT NULL, 
	[2]			Text (255) NOT NULL, 
	[3]			Text (255) NOT NULL, 
	[4]			Text (255) NOT NULL, 
	[5]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL
);

CREATE TABLE [Q_S_QA_DIM_122_C]
 (
	[用途名]			Text (255) NOT NULL, 
	[製造材質]			Text (255) NOT NULL, 
	[調質]			Text (255) NOT NULL, 
	[用途コード]			Text (255) NOT NULL, 
	[納入先名]			Text (255) NOT NULL, 
	[板厚下]			Text (255) NOT NULL, 
	[板厚上]			Text (255) NOT NULL, 
	[オーダー板厚T]			Text (255) NOT NULL, 
	[幅公差]			Text (255) NOT NULL, 
	[丈公差]			Text (255) NOT NULL, 
	[オーダー板幅]			Text (255) NOT NULL, 
	[パターン]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL
);

CREATE TABLE [ラテフリ直角]
 (
	[管理番号]			Text (255) NOT NULL, 
	[納入先]			Text (255) NOT NULL, 
	[納入先コード]			Text (255) NOT NULL, 
	[取引先]			Text (255) NOT NULL, 
	[取引先コード]			Text (255) NOT NULL, 
	[用途コード]			Text (255) NOT NULL, 
	[製造材質]			Text (255) NOT NULL, 
	[製造調質]			Text (255) NOT NULL, 
	[厚下]			Text (255) NOT NULL, 
	[厚上]			Text (255) NOT NULL, 
	[製造板厚]			Text (255) NOT NULL, 
	[幅下]			Text (255) NOT NULL, 
	[幅上]			Text (255) NOT NULL, 
	[製造板幅]			Text (255) NOT NULL, 
	[丈下]			Text (255) NOT NULL, 
	[丈上]			Text (255) NOT NULL, 
	[製造板丈]			Text (255) NOT NULL, 
	[フリー反り]			Text (255) NOT NULL, 
	[ラテラルボー]			Text (255) NOT NULL, 
	[直角度]			Text (255) NOT NULL, 
	[公差]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL
);

CREATE TABLE [Q_S_QA_124_別表1]
 (
	[管理番号]			Text (255) NOT NULL, 
	[欠陥名]			Text (255) NOT NULL, 
	[1]			Text (255) NOT NULL, 
	[2]			Text (255) NOT NULL, 
	[3]			Text (255) NOT NULL, 
	[4]			Text (255) NOT NULL, 
	[5]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL, 
	[更新日]			Text (255) NOT NULL
);

CREATE TABLE [Q_S_KE_ALL_053]
 (
	[材質]			Text (255) NOT NULL, 
	[調質]			Text (255) NOT NULL, 
	[バーコール下限値]			Text (255) NOT NULL, 
	[引張強さ(Mpa)]			Text (255) NOT NULL
);

CREATE TABLE [中歪3級]
 (
	[平坦度上限値参照元]			Text (255) NOT NULL, 
	[用途コード]			Text (255) NOT NULL, 
	[板厚下]			Text (255) NOT NULL, 
	[板厚上]			Text (255) NOT NULL, 
	[500]			Text (255) NOT NULL, 
	[4000]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL, 
	[管理番号]			Text (255)
);

CREATE TABLE [耳歪3級]
 (
	[平坦度上限値参照元]			Text (255) NOT NULL, 
	[用途コード]			Text (255) NOT NULL, 
	[板厚下]			Text (255) NOT NULL, 
	[板厚上]			Text (255) NOT NULL, 
	[500]			Text (255) NOT NULL, 
	[4000]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL, 
	[管理番号]			Text (255)
);

CREATE TABLE [耳歪4級]
 (
	[平坦度上限値参照元]			Text (255) NOT NULL, 
	[用途コード]			Text (255) NOT NULL, 
	[板厚下]			Text (255) NOT NULL, 
	[板厚上]			Text (255) NOT NULL, 
	[500]			Text (255) NOT NULL, 
	[4000]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL, 
	[管理番号]			Text (255)
);

CREATE TABLE [~TMPCLP531581]
 (
	[材質]			Text (255) NOT NULL, 
	[用途コード]			Text (255) NOT NULL, 
	[板厚下]			Text (255) NOT NULL, 
	[板厚上]			Text (255) NOT NULL, 
	[厚]			Text (255) NOT NULL, 
	[飛込み]			Text (255) NOT NULL, 
	[キズ]			Text (255) NOT NULL, 
	[鋳造剥離AL圧込]			Text (255) NOT NULL, 
	[AL付着跡]			Text (255) NOT NULL, 
	[汚れボタ落ち]			Text (255) NOT NULL, 
	[カミコミ跡ショック跡]			Text (255) NOT NULL, 
	[ナメカワフラグ]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL
);

CREATE TABLE [Q_S_QA_ALL_129]
 (
	[材質]			Text (255) NOT NULL, 
	[調質]			Text (255) NOT NULL, 
	[用途コード]			Text (255) NOT NULL, 
	[板厚下]			Text (255) NOT NULL, 
	[板厚上]			Text (255) NOT NULL, 
	[厚]			Text (255) NOT NULL, 
	[500]			Text (255) NOT NULL, 
	[750]			Text (255) NOT NULL, 
	[1000]			Text (255) NOT NULL, 
	[1250]			Text (255) NOT NULL, 
	[1525]			Text (255) NOT NULL, 
	[2000]			Text (255) NOT NULL, 
	[2500]			Text (255) NOT NULL, 
	[3050]			Text (255) NOT NULL, 
	[4500]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL
);

CREATE TABLE [Q_S_QA_DIM_122_B]
 (
	[材質]			Text (255) NOT NULL, 
	[調質]			Text (255) NOT NULL, 
	[用途コード]			Text (255) NOT NULL, 
	[板厚下]			Text (255) NOT NULL, 
	[板厚上]			Text (255) NOT NULL, 
	[厚]			Text (255) NOT NULL, 
	[板幅下]			Text (255) NOT NULL, 
	[板幅上]			Text (255) NOT NULL, 
	[幅]			Text (255) NOT NULL, 
	[板幅]			Text (255) NOT NULL, 
	[板丈]			Text (255) NOT NULL, 
	[中歪・耳歪]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL
);

CREATE TABLE [Q_S_QA_DIM_122_D]
 (
	[用途名]			Text (255) NOT NULL, 
	[製造材質]			Text (255) NOT NULL, 
	[製造調質]			Text (255) NOT NULL, 
	[用途コード]			Text (255) NOT NULL, 
	[納入先名]			Text (255) NOT NULL, 
	[オーダー板厚]			Text (255) NOT NULL, 
	[幅公差]			Text (255) NOT NULL, 
	[丈公差]			Text (255) NOT NULL, 
	[中歪耳歪]			Text (255) NOT NULL, 
	[フラットネス品質等級]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL
);

CREATE TABLE [Q_S_QA_DIM_122_A]
 (
	[平坦度上限値参照元]			Text (255) NOT NULL, 
	[材質]			Text (255) NOT NULL, 
	[調質]			Text (255) NOT NULL, 
	[用途コード]			Text (255) NOT NULL, 
	[板厚下]			Text (255) NOT NULL, 
	[板厚上]			Text (255) NOT NULL, 
	[厚]			Text (255) NOT NULL, 
	[1m当たり平坦度上限]			Text (255) NOT NULL, 
	[500]			Text (255) NOT NULL, 
	[750]			Text (255) NOT NULL, 
	[1004]			Text (255) NOT NULL, 
	[1254]			Text (255) NOT NULL, 
	[1529]			Text (255) NOT NULL, 
	[2004]			Text (255) NOT NULL, 
	[2504]			Text (255) NOT NULL, 
	[3054]			Text (255) NOT NULL, 
	[4000]			Text (255) NOT NULL, 
	[納入フラグ]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL
);

CREATE TABLE [~TMPCLP608031]
 (
	[番号]			Text (255) NOT NULL, 
	[フラグ]			Text (255) NOT NULL, 
	[平坦度上限値参照元]			Text (255) NOT NULL, 
	[用途名]			Text (255) NOT NULL, 
	[用途コード]			Text (255) NOT NULL, 
	[製造材質]			Text (255) NOT NULL, 
	[オーダー材質]			Text (255) NOT NULL, 
	[製造調質]			Text (255) NOT NULL, 
	[オーダー調質]			Text (255) NOT NULL, 
	[板厚下]			Text (255) NOT NULL, 
	[板厚上]			Text (255) NOT NULL, 
	[オーダー板厚]			Text (255) NOT NULL, 
	[板幅下]			Text (255) NOT NULL, 
	[板幅上]			Text (255) NOT NULL, 
	[オーダー板幅]			Text (255) NOT NULL, 
	[オーダー板丈]			Text (255) NOT NULL, 
	[納入先]			Text (255) NOT NULL, 
	[納入先コード]			Text (255) NOT NULL, 
	[巾取数幅分割数のこと]			Text (255) NOT NULL, 
	[検査番号の頭]			Text (255) NOT NULL, 
	[形状コード]			Text (255) NOT NULL, 
	[1m当たり平坦度上限]			Text (255) NOT NULL, 
	[H500]			Text (255) NOT NULL, 
	[H750]			Text (255) NOT NULL, 
	[H1004]			Text (255) NOT NULL, 
	[H1254]			Text (255) NOT NULL, 
	[H1529]			Text (255) NOT NULL, 
	[H2004]			Text (255) NOT NULL, 
	[H2504]			Text (255) NOT NULL, 
	[H3054]			Text (255) NOT NULL, 
	[T500]			Text (255) NOT NULL, 
	[T750]			Text (255) NOT NULL, 
	[T1004]			Text (255) NOT NULL, 
	[T1254]			Text (255) NOT NULL, 
	[T1529]			Text (255) NOT NULL, 
	[T2004]			Text (255) NOT NULL, 
	[T2504]			Text (255) NOT NULL, 
	[T3054]			Text (255) NOT NULL, 
	[中歪・耳歪]			Text (255) NOT NULL, 
	[表示コメント]			Text (255) NOT NULL, 
	[コメント(原文)]			Text (255) NOT NULL, 
	[メモ]			Text (255) NOT NULL
);

CREATE TABLE [~TMPCLP338731]
 (
	[板厚下]			Text (255) NOT NULL, 
	[板厚上]			Text (255) NOT NULL, 
	[H500]			Text (255) NOT NULL, 
	[H1000]			Text (255) NOT NULL, 
	[H1500]			Text (255) NOT NULL, 
	[H1530]			Text (255) NOT NULL, 
	[H2000]			Text (255) NOT NULL, 
	[H9999]			Text (255) NOT NULL, 
	[T500]			Text (255) NOT NULL, 
	[T1000]			Text (255) NOT NULL, 
	[T1500]			Text (255) NOT NULL, 
	[T1530]			Text (255) NOT NULL, 
	[T1999]			Text (255) NOT NULL, 
	[T2000]			Text (255) NOT NULL, 
	[T9999]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL
);

CREATE TABLE [~TMPCLP262511]
 (
	[板厚下]			Text (255) NOT NULL, 
	[板厚上]			Text (255) NOT NULL, 
	[板丈]			Text (255) NOT NULL, 
	[H500]			Text (255) NOT NULL, 
	[H1000]			Text (255) NOT NULL, 
	[H1500]			Text (255) NOT NULL, 
	[H1530]			Text (255) NOT NULL, 
	[H2000]			Text (255) NOT NULL, 
	[H9999]			Text (255) NOT NULL, 
	[T500]			Text (255) NOT NULL, 
	[T1000]			Text (255) NOT NULL, 
	[T1500]			Text (255) NOT NULL, 
	[T1530]			Text (255) NOT NULL, 
	[T1999]			Text (255) NOT NULL, 
	[T2000]			Text (255) NOT NULL, 
	[T9999]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL
);

CREATE TABLE [~TMPCLP412091]
 (
	[板厚下]			Text (255) NOT NULL, 
	[板厚上]			Text (255) NOT NULL, 
	[板丈下]			Text (255) NOT NULL, 
	[板丈上]			Text (255) NOT NULL, 
	[板丈]			Text (255) NOT NULL, 
	[H500]			Text (255) NOT NULL, 
	[H1000]			Text (255) NOT NULL, 
	[H1500]			Text (255) NOT NULL, 
	[H1530]			Text (255) NOT NULL, 
	[H2000]			Text (255) NOT NULL, 
	[H9999]			Text (255) NOT NULL, 
	[T500]			Text (255) NOT NULL, 
	[T1000]			Text (255) NOT NULL, 
	[T1500]			Text (255) NOT NULL, 
	[T1530]			Text (255) NOT NULL, 
	[T1999]			Text (255) NOT NULL, 
	[T2000]			Text (255) NOT NULL, 
	[T9999]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL
);

CREATE TABLE [~TMPCLP379751]
 (
	[板厚下]			Text (255) NOT NULL, 
	[板厚上]			Text (255) NOT NULL, 
	[板丈下]			Text (255) NOT NULL, 
	[板丈上]			Text (255) NOT NULL, 
	[板丈]			Text (255) NOT NULL, 
	[H500]			Text (255) NOT NULL, 
	[H1000]			Text (255) NOT NULL, 
	[H1500]			Text (255) NOT NULL, 
	[H1530]			Text (255) NOT NULL, 
	[H2000]			Text (255) NOT NULL, 
	[H9999]			Text (255) NOT NULL, 
	[T500]			Text (255) NOT NULL, 
	[T1000]			Text (255) NOT NULL, 
	[T1500]			Text (255) NOT NULL, 
	[T1530]			Text (255) NOT NULL, 
	[T1999]			Text (255) NOT NULL, 
	[T2000]			Text (255) NOT NULL, 
	[T9999]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL
);

CREATE TABLE [ソリ4級]
 (
	[板厚下]			Text (255) NOT NULL, 
	[板厚上]			Text (255) NOT NULL, 
	[H500]			Text (255) NOT NULL, 
	[H1000]			Text (255) NOT NULL, 
	[H1500]			Text (255) NOT NULL, 
	[H1530]			Text (255) NOT NULL, 
	[H2000]			Text (255) NOT NULL, 
	[H9999]			Text (255) NOT NULL, 
	[T500]			Text (255) NOT NULL, 
	[T1000]			Text (255) NOT NULL, 
	[T1500]			Text (255) NOT NULL, 
	[T1530]			Text (255) NOT NULL, 
	[T2000]			Text (255) NOT NULL, 
	[T9999]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL, 
	[管理番号]			Text (255), 
	[T1999]			Text (255)
);

CREATE TABLE [ソリ3級]
 (
	[板厚下]			Text (255) NOT NULL, 
	[板厚上]			Text (255) NOT NULL, 
	[H500]			Text (255) NOT NULL, 
	[H1000]			Text (255) NOT NULL, 
	[H1500]			Text (255) NOT NULL, 
	[H1530]			Text (255) NOT NULL, 
	[H2000]			Text (255) NOT NULL, 
	[H9999]			Text (255) NOT NULL, 
	[T500]			Text (255) NOT NULL, 
	[T1000]			Text (255) NOT NULL, 
	[T1500]			Text (255) NOT NULL, 
	[T1530]			Text (255) NOT NULL, 
	[T2000]			Text (255) NOT NULL, 
	[T9999]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL, 
	[管理番号]			Text (255), 
	[T1999]			Text (255)
);

CREATE TABLE [~TMPCLP559851]
 (
	[番号]			Text (255) NOT NULL, 
	[フラグ]			Text (255) NOT NULL, 
	[平坦度上限値参照元]			Text (255) NOT NULL, 
	[用途名]			Text (255) NOT NULL, 
	[用途コード]			Text (255) NOT NULL, 
	[製造材質]			Text (255) NOT NULL, 
	[オーダー材質]			Text (255) NOT NULL, 
	[製造調質]			Text (255) NOT NULL, 
	[オーダー調質]			Text (255) NOT NULL, 
	[板厚下]			Text (255) NOT NULL, 
	[板厚上]			Text (255) NOT NULL, 
	[オーダー板厚]			Text (255) NOT NULL, 
	[板幅下]			Text (255) NOT NULL, 
	[板幅上]			Text (255) NOT NULL, 
	[オーダー板幅]			Text (255) NOT NULL, 
	[オーダー板丈]			Text (255) NOT NULL, 
	[納入先]			Text (255) NOT NULL, 
	[納入先コード]			Text (255) NOT NULL, 
	[巾取数幅分割数のこと]			Text (255) NOT NULL, 
	[検査番号の頭]			Text (255) NOT NULL, 
	[形状コード]			Text (255) NOT NULL, 
	[1m当たり平坦度上限]			Text (255) NOT NULL, 
	[H500]			Text (255) NOT NULL, 
	[H750]			Text (255) NOT NULL, 
	[H1004]			Text (255) NOT NULL, 
	[H1254]			Text (255) NOT NULL, 
	[H1529]			Text (255) NOT NULL, 
	[H2004]			Text (255) NOT NULL, 
	[H2504]			Text (255) NOT NULL, 
	[H3054]			Text (255) NOT NULL, 
	[T500]			Text (255) NOT NULL, 
	[T750]			Text (255) NOT NULL, 
	[T1004]			Text (255) NOT NULL, 
	[T1254]			Text (255) NOT NULL, 
	[T1529]			Text (255) NOT NULL, 
	[T2004]			Text (255) NOT NULL, 
	[T2504]			Text (255) NOT NULL, 
	[T3054]			Text (255) NOT NULL, 
	[中歪・耳歪]			Text (255) NOT NULL, 
	[表示コメント]			Text (255) NOT NULL, 
	[コメント(原文)]			Text (255) NOT NULL, 
	[メモ]			Text (255) NOT NULL
);

CREATE TABLE [~TMPCLP411111]
 (
	[番号]			Text (255) NOT NULL, 
	[フラグ]			Text (255) NOT NULL, 
	[平坦度上限値参照元]			Text (255) NOT NULL, 
	[用途名]			Text (255) NOT NULL, 
	[用途コード]			Text (255) NOT NULL, 
	[製造材質]			Text (255) NOT NULL, 
	[オーダー材質]			Text (255) NOT NULL, 
	[製造調質]			Text (255) NOT NULL, 
	[オーダー調質]			Text (255) NOT NULL, 
	[板厚下]			Text (255) NOT NULL, 
	[板厚上]			Text (255) NOT NULL, 
	[オーダー板厚]			Text (255) NOT NULL, 
	[板幅下]			Text (255) NOT NULL, 
	[板幅上]			Text (255) NOT NULL, 
	[オーダー板幅]			Text (255) NOT NULL, 
	[オーダー板丈]			Text (255) NOT NULL, 
	[納入先]			Text (255) NOT NULL, 
	[納入先コード]			Text (255) NOT NULL, 
	[巾取数幅分割数のこと]			Text (255) NOT NULL, 
	[検査番号の頭]			Text (255) NOT NULL, 
	[形状コード]			Text (255) NOT NULL, 
	[1m当たり平坦度上限]			Text (255) NOT NULL, 
	[H500]			Text (255) NOT NULL, 
	[H750]			Text (255) NOT NULL, 
	[H1004]			Text (255) NOT NULL, 
	[H1254]			Text (255) NOT NULL, 
	[H1529]			Text (255) NOT NULL, 
	[H2004]			Text (255) NOT NULL, 
	[H2504]			Text (255) NOT NULL, 
	[H3054]			Text (255) NOT NULL, 
	[T500]			Text (255) NOT NULL, 
	[T750]			Text (255) NOT NULL, 
	[T1004]			Text (255) NOT NULL, 
	[T1254]			Text (255) NOT NULL, 
	[T1529]			Text (255) NOT NULL, 
	[T2004]			Text (255) NOT NULL, 
	[T2504]			Text (255) NOT NULL, 
	[T3054]			Text (255) NOT NULL, 
	[中歪・耳歪]			Text (255) NOT NULL, 
	[表示コメント]			Text (255) NOT NULL, 
	[コメント(原文)]			Text (255) NOT NULL, 
	[メモ]			Text (255) NOT NULL
);

CREATE TABLE [~TMPCLP144201]
 (
	[欠陥名]			Text (255) NOT NULL, 
	[1]			Text (255) NOT NULL, 
	[2]			Text (255) NOT NULL, 
	[3]			Text (255) NOT NULL, 
	[4]			Text (255) NOT NULL, 
	[5]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL
);

CREATE TABLE [Q_S_QA_SRF_128_別表2]
 (
	[管理番号]			Text (255) NOT NULL, 
	[パターン]			Text (255) NOT NULL, 
	[条件]			Text (255) NOT NULL, 
	[トビコミ]			Text (255) NOT NULL, 
	[キズ,スリキズ]			Text (255) NOT NULL, 
	[突上げ,突下げ]			Text (255) NOT NULL, 
	[共ズレ]			Text (255) NOT NULL, 
	[汚れ,汚れ圧込み]			Text (255) NOT NULL, 
	[油寄り,レベラースジ]			Text (255) NOT NULL, 
	[ロールキズ]			Text (255) NOT NULL, 
	[ロールマーク]			Text (255) NOT NULL, 
	[噛み込み跡]			Text (255) NOT NULL, 
	[切断面異常]			Text (255) NOT NULL, 
	[エッジ変色,油残り跡,油焼け]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL, 
	[更新日]			Text (255) NOT NULL
);

CREATE TABLE [Q_S_QA_DIM_179_白銅滑川]
 (
	[管理番号]			Text (255) NOT NULL, 
	[材質]			Text (255) NOT NULL, 
	[調質]			Text (255) NOT NULL, 
	[用途コード]			Text (255) NOT NULL, 
	[板厚下]			Text (255) NOT NULL, 
	[板厚上]			Text (255) NOT NULL, 
	[厚]			Text (255) NOT NULL, 
	[板幅下]			Text (255) NOT NULL, 
	[板幅上]			Text (255) NOT NULL, 
	[幅]			Text (255) NOT NULL, 
	[1500]			Text (255) NOT NULL, 
	[2500]			Text (255) NOT NULL, 
	[9999]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL
);

CREATE TABLE [Q_S_QA_DIM_179_日車]
 (
	[管理番号]			Text (255) NOT NULL, 
	[材質]			Text (255) NOT NULL, 
	[調質]			Text (255) NOT NULL, 
	[用途コード]			Text (255) NOT NULL, 
	[板厚下]			Text (255) NOT NULL, 
	[板厚上]			Text (255) NOT NULL, 
	[厚]			Text (255) NOT NULL, 
	[板幅下]			Text (255) NOT NULL, 
	[板幅上]			Text (255) NOT NULL, 
	[幅]			Text (255) NOT NULL, 
	[750]			Text (255) NOT NULL, 
	[1525]			Text (255) NOT NULL, 
	[2300]			Text (255) NOT NULL, 
	[3050]			Text (255) NOT NULL, 
	[4500]			Text (255) NOT NULL, 
	[5200]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL
);

CREATE TABLE [Q_S_QA_SRF_128_別表1]
 (
	[管理番号]			Text (255) NOT NULL, 
	[表面処理コード]			Text (255) NOT NULL, 
	[材質]			Text (255) NOT NULL, 
	[用途名]			Text (255) NOT NULL, 
	[用途コード]			Text (255) NOT NULL, 
	[厚フラグ]			Text (255) NOT NULL, 
	[板厚1]			Text (255) NOT NULL, 
	[パターン1]			Text (255) NOT NULL, 
	[パターン2]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL, 
	[更新日]			Text (255) NOT NULL
);

CREATE TABLE [Q_S_QA_124_別表2]
 (
	[管理番号]			Text (255) NOT NULL, 
	[材質]			Text (255) NOT NULL, 
	[用途コード]			Text (255) NOT NULL, 
	[板厚下]			Text (255) NOT NULL, 
	[板厚上]			Text (255) NOT NULL, 
	[厚]			Text (255) NOT NULL, 
	[飛込み]			Text (255) NOT NULL, 
	[キズ/スリキズ/スリ跡]			Text (255) NOT NULL, 
	[鋳造剥離AL圧込]			Text (255) NOT NULL, 
	[AL付着跡]			Text (255) NOT NULL, 
	[汚れボタ落ち]			Text (255) NOT NULL, 
	[カミコミ跡ショック跡]			Text (255) NOT NULL, 
	[備考]			Text (255) NOT NULL, 
	[更新日]			Text (255) NOT NULL, 
	[板幅下]			Text (255) NOT NULL, 
	[板幅上]			Text (255) NOT NULL, 
	[幅]			Text (255) NOT NULL, 
	[板丈下]			Text (255) NOT NULL, 
	[板丈上]			Text (255) NOT NULL, 
	[丈]			Text (255) NOT NULL, 
	[納入先]			Text (255) NOT NULL
);

CREATE TABLE [平坦度一覧QP追加]
 (
	[フラグ]			Text (255) NOT NULL, 
	[平坦度上限値参照元]			Text (255) NOT NULL, 
	[用途名]			Text (255) NOT NULL, 
	[用途コード]			Text (255) NOT NULL, 
	[製造材質]			Text (255) NOT NULL, 
	[オーダー材質]			Text (255) NOT NULL, 
	[製造調質]			Text (255) NOT NULL, 
	[オーダー調質]			Text (255) NOT NULL, 
	[板厚下]			Text (255) NOT NULL, 
	[板厚上]			Text (255) NOT NULL, 
	[オーダー板厚]			Text (255) NOT NULL, 
	[板幅下]			Text (255) NOT NULL, 
	[板幅上]			Text (255) NOT NULL, 
	[オーダー板幅]			Text (255) NOT NULL, 
	[板丈下]			Text (255) NOT NULL, 
	[板丈上]			Text (255) NOT NULL, 
	[オーダー板丈]			Text (255) NOT NULL, 
	[納入先]			Text (255) NOT NULL, 
	[送り先]			Text (255) NOT NULL, 
	[送り先コード]			Text (255) NOT NULL, 
	[納入先コード]			Text (255) NOT NULL, 
	[取引先]			Text (255) NOT NULL, 
	[取引先コード]			Text (255) NOT NULL, 
	[巾取数幅分割数のこと]			Text (255) NOT NULL, 
	[検査番号の頭]			Text (255) NOT NULL, 
	[形状コード]			Text (255) NOT NULL, 
	[1m当たり平坦度上限]			Text (255) NOT NULL, 
	[H500]			Text (255) NOT NULL, 
	[H750]			Text (255) NOT NULL, 
	[H1004]			Text (255) NOT NULL, 
	[H1254]			Text (255) NOT NULL, 
	[H1750]			Text (255) NOT NULL, 
	[H2004]			Text (255) NOT NULL, 
	[H2504]			Text (255) NOT NULL, 
	[H3054]			Text (255) NOT NULL, 
	[T500]			Text (255) NOT NULL, 
	[T750]			Text (255) NOT NULL, 
	[T1004]			Text (255) NOT NULL, 
	[T1254]			Text (255) NOT NULL, 
	[T1750]			Text (255) NOT NULL, 
	[T2004]			Text (255) NOT NULL, 
	[T2504]			Text (255) NOT NULL, 
	[T3054]			Text (255) NOT NULL, 
	[表示コメント]			Text (255) NOT NULL, 
	[中歪・耳歪]			Text (255) NOT NULL, 
	[更新日]			Text (255) NOT NULL, 
	[管理番号]			Long Integer NOT NULL
);

CREATE TABLE [欠陥名]
 (
	[管理番号]			Text (255) NOT NULL, 
	[欠陥名]			Text (255) NOT NULL
);

CREATE TABLE [製品規格直角度]
 (
	[管理番号]			Long Integer, 
	[等級]			Text (10) NOT NULL, 
	[巾_下限]			Double NOT NULL, 
	[巾_上限]			Double NOT NULL, 
	[丈_下限]			Double NOT NULL, 
	[丈_上限]			Double NOT NULL, 
	[公差]			Double NOT NULL, 
	[備考]			Text (255)
);

CREATE TABLE [製品規格ラテラルボー]
 (
	[管理番号]			Long Integer, 
	[等級]			Text (10) NOT NULL, 
	[厚_下限]			Double NOT NULL, 
	[厚_上限]			Double NOT NULL, 
	[巾_下限]			Double NOT NULL, 
	[巾_上限]			Double NOT NULL, 
	[丈_下限]			Double NOT NULL, 
	[丈_上限]			Double NOT NULL, 
	[公差]			Double NOT NULL, 
	[備考]			Text (255)
);

CREATE TABLE [外注条件]
 (
	[管理番号]			Long Integer NOT NULL, 
	[条件名]			Text (100) NOT NULL, 
	[戻り値]			Text (50) NOT NULL, 
	[優先順位]			Long Integer NOT NULL, 
	[有効]			Boolean NOT NULL, 
	[備考]			Text (255)
);

CREATE TABLE [外注条件明細]
 (
	[管理番号]			Long Integer, 
	[条件管理番号]			Long Integer NOT NULL, 
	[条件番号]			Long Integer NOT NULL, 
	[対象項目]			Text (50) NOT NULL, 
	[判定方法]			Text (10) NOT NULL, 
	[判定値]			Text (100) NOT NULL, 
	[有効]			Boolean NOT NULL, 
	[備考]			Text (255)
);

CREATE TABLE [ゲージ判定条件]
 (
	[管理番号]			Long Integer NOT NULL, 
	[重量_超]			Double, 
	[丈_以上]			Double, 
	[丈_未満]			Double, 
	[丈_超]			Double, 
	[幅_未満]			Double, 
	[材質]			Text (50), 
	[調質]			Text (50), 
	[コース含む]			Text (50), 
	[有効]			Boolean NOT NULL, 
	[説明]			Text (255)
);



<%
Function getWarmreceptionMaster()
	Dim aData(16)	'資格・検定名で選ぶディクショナリ配列
	Dim idx

	idx = -1

	'<テンプレート>
	'idx = idx + 1
	'Set aData(idx) = Server.CreateObject("scripting.dictionary")
	'aData(idx).Add "Category","license"
	'aData(idx).Add "ID",""
	'aData(idx).Add "優先",""
	'aData(idx).Add "種別",""
	'aData(idx).Add "名称",""
	'aData(idx).Add "概要",""
	'aData(idx).Add "詳細",""
	'aData(idx).Add "費用",""
	'aData(idx).Add "合格率",""
	'aData(idx).Add "団体名",""
	'aData(idx).Add "教育機関",""
	'aData(idx).Add "講座内容",""
	'aData(idx).Add "価格",""
	'aData(idx).Add "特典",""
	'aData(idx).Add "クーポン",""
	'aData(idx).Add "対象業種",""
	'aData(idx).Add "価格2",""
	'aData(idx).Add "概要2",""
	'aData(idx).Add "特典2",""
	'</テンプレート>

	'<秘書検定認定試験1級>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","license"
	aData(idx).Add "ID","0401"
	aData(idx).Add "優先",""
	aData(idx).Add "種別","ビジネス(資格)系 ／ 通信教育"
	aData(idx).Add "名称","秘書検定認定試験（1級）"
	aData(idx).Add "名称TITLE","秘書検定認定試験（1級）"
	aData(idx).Add "概要","【全業種対応の資格】" & vbCrLf & vbCrLf & "『文部科学省認定の公的資格』として、秘書のみならず事務職全般の高いスキルを証明できる資格"
	aData(idx).Add "詳細","秘書という特別な職務だけでなく、合格することで一般事務職の知識と技能をもっている証明として注目されています。事務処理・情報処理・接遇のエキスパートとして、あらゆる組織で活躍できる場を広げることができます。"
	aData(idx).Add "費用","受験料は" & vbCrLf & "1級　6,000円" & vbCrLf & "準1級　4,800円" & vbCrLf & "2級　3,700円" & vbCrlf & "3級　2,500円"
	aData(idx).Add "合格率","1級　25．3％" & vbCrlf & "準1級　29．8％" & vbCrlf & "2級　43．5％" & vbCrlf & "3級　64．0％" & vbCrlf & "（08年11月）"
	aData(idx).Add "団体名","財団法人実務技能検定協会"
	aData(idx).Add "教育機関","産業能率大学総合研究所"
	aData(idx).Add "講座内容","★通信教育" & vbCrLf & "（ﾃｷｽﾄ・添削有り）"
	aData(idx).Add "価格","1級　26,250円" & vbCrLf & "準1級　26,250円"
	aData(idx).Add "特典","受講料20％OFF" & vbCrLf & "※支払方法：代金引き換えによる受講となります。"
	aData(idx).Add "クーポン",""
	aData(idx).Add "対象業種","全業種対応の資格"
	aData(idx).Add "価格2","1級　26,250円<br>準1級　26,250円"
	aData(idx).Add "概要2","『文部科学省認定の公的資格』として、秘書のみならず事務職全般の高いスキルを証明できる資格"
	aData(idx).Add "特典2","受講料<br>20％OFF"
	'</秘書検定認定試験1級>

	'<秘書検定認定試験2級>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","license"
	aData(idx).Add "ID","0402"
	aData(idx).Add "優先",""
	aData(idx).Add "種別","ビジネス(資格)系 ／ 通信教育"
	aData(idx).Add "名称","秘書検定認定試験（2級）"
	aData(idx).Add "名称TITLE","秘書検定認定試験（2級）"
	aData(idx).Add "概要","【全業種対応の資格】" & vbCrLf & vbCrLf & "『文部科学省認定の公的資格』として、秘書のみならず事務職全般の高いスキルを証明できる資格"
	aData(idx).Add "詳細","秘書という特別な職務だけでなく、合格することで一般事務職の知識と技能をもっている証明として注目されています。事務処理・情報処理・接遇のエキスパートとして、あらゆる組織で活躍できる場を広げることができます。"
	aData(idx).Add "費用","受験料は" & vbCrLf & "1級　6,000円" & vbCrLf & "準1級　4,800円" & vbCrLf & "2級　3,700円" & vbCrlf & "3級　2,500円"
	aData(idx).Add "合格率","1級　25．3％" & vbCrlf & "準1級　29．8％" & vbCrlf & "2級　43．5％" & vbCrlf & "3級　64．0％" & vbCrlf & "（08年11月）"
	aData(idx).Add "団体名","財団法人実務技能検定協会"
	aData(idx).Add "教育機関","産業能率大学総合研究所"
	aData(idx).Add "講座内容","★通信教育" & vbCrLf & "（ﾃｷｽﾄ・添削有り）"
	aData(idx).Add "価格","2級　21,000円"
	aData(idx).Add "特典","受講料20％OFF" & vbCrLf & "※支払方法：代金引き換えによる受講となります。"
	aData(idx).Add "クーポン",""
	aData(idx).Add "対象業種","全業種対応の資格"
	aData(idx).Add "価格2","2級　21,000円"
	aData(idx).Add "概要2","文部科学省認定の公的資格』として、秘書のみならず事務職全般の高いスキルを証明できる資格"
	aData(idx).Add "特典2","受講料<br>20％OFF"

	'</秘書検定認定試験2級>

	'<秘書検定認定試験3級>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","license"
	aData(idx).Add "ID","0403"
	aData(idx).Add "優先",""
	aData(idx).Add "種別","ビジネス(資格)系 ／ 通信教育"
	aData(idx).Add "名称","秘書検定認定試験（3級）"
	aData(idx).Add "名称TITLE","秘書検定認定試験（3級）"
	aData(idx).Add "概要","【全業種対応の資格】" & vbCrLf & vbCrLf & "『文部科学省認定の公的資格』として、秘書のみならず事務職全般の高いスキルを証明できる資格"
	aData(idx).Add "詳細","秘書という特別な職務だけでなく、合格することで一般事務職の知識と技能をもっている証明として注目されています。事務処理・情報処理・接遇のエキスパートとして、あらゆる組織で活躍できる場を広げることができます。"
	aData(idx).Add "費用","受験料は" & vbCrLf & "1級　6,000円" & vbCrLf & "準1級　4,800円" & vbCrLf & "2級　3,700円" & vbCrlf & "3級　2,500円"
	aData(idx).Add "合格率","1級　25．3％" & vbCrlf & "準1級　29．8％" & vbCrlf & "2級　43．5％" & vbCrlf & "3級　64．0％" & vbCrlf & "（08年11月）"
	aData(idx).Add "団体名","財団法人実務技能検定協会"
	aData(idx).Add "教育機関","産業能率大学総合研究所"
	aData(idx).Add "講座内容","★通信教育" & vbCrLf & "（ﾃｷｽﾄ・添削有り）"
	aData(idx).Add "価格","3級　25,200円"
	aData(idx).Add "特典","受講料20％OFF" & vbCrLf & "※支払方法：代金引き換えによる受講となります。"
	aData(idx).Add "クーポン",""
	aData(idx).Add "対象業種","全業種対応の資格"
	aData(idx).Add "価格2","3級　25,200円"
	aData(idx).Add "概要2","『文部科学省認定の公的資格』として、秘書のみならず事務職全般の高いスキルを証明できる資格"
	aData(idx).Add "特典2","受講料<br>20％OFF"

	'</秘書検定認定試験3級>

	'<ﾒﾝﾀﾙﾍﾙｽ・ﾏﾈｼﾞﾒﾝﾄ検定Ⅰ種>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","license"
	aData(idx).Add "ID","0501"
	aData(idx).Add "優先",""
	aData(idx).Add "種別","ビジネス(資格)系 ／ 通信教育"
	aData(idx).Add "名称","メンタルヘルス・マネジメント検定（Ⅰ種）"
	aData(idx).Add "名称TITLE","ﾒﾝﾀﾙﾍﾙｽ・ﾏﾈｼﾞﾒﾝﾄ検定（Ⅰ種）"
	aData(idx).Add "概要","【総務・人事・労務・事務の方にお薦めの資格】" & vbCrLf & vbCrLf & "企業内での適切なメンタルヘルス対策を講じる総務・人事労務管理に携わる方にお勧めの検定"
	aData(idx).Add "詳細","心の病を抱える労働者の増加が社会問題化しており、働く人の「心の健康管理」に関心が高まっています。厚生労働省では職場における適切かつ有効なメンタルヘルス対策の実施を推進しています。本検定試験は、厚生労働省の「労働者の心の健康の保持増進のための指針」に基づいて構築されています。"
	aData(idx).Add "費用","Ⅰ種　10,500円" & vbCrLf & "Ⅱ種　　6,300円" & vbCrlf & "Ⅲ種　　4,200円"
	aData(idx).Add "合格率","08年度" & vbCrLf & "Ⅰ種　11．1％" & vbCrLf & "Ⅱ種　70．7％" & vbCrLf & "Ⅲ種　87．4％"
	aData(idx).Add "団体名","大阪商工会議所"
	aData(idx).Add "教育機関","産業能率大学総合研究所"
	aData(idx).Add "講座内容","★通信教育" & vbCrLf & "（ﾃｷｽﾄ・添削有り）"
	aData(idx).Add "価格","受講料" & vbCrLf & "Ⅰ種　24,1500円"
	aData(idx).Add "特典","特別受講料" & vbCrLf & "Ⅰ種21,000円" & vbCrLf & "※支払方法：代金引き換えによる受講となります。"
	aData(idx).Add "クーポン",""
	aData(idx).Add "対象業種","総務・人事・労務・事務の方にお薦め"
	aData(idx).Add "価格2","Ⅰ種　24,1500円"
	aData(idx).Add "概要2","企業内での適切なメンタルヘルス対策を講じる総務・人事労務管理に携わる方にお勧めの検定"
	aData(idx).Add "特典2","特別受講料<br>Ⅰ種21,000円"

	'</ﾒﾝﾀﾙﾍﾙｽ・ﾏﾈｼﾞﾒﾝﾄ検定Ⅰ種>

	'<ﾒﾝﾀﾙﾍﾙｽ・ﾏﾈｼﾞﾒﾝﾄ検定Ⅱ種>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","license"
	aData(idx).Add "ID","0502"
	aData(idx).Add "優先",""
	aData(idx).Add "種別","ビジネス(資格)系 ／ 通信教育"
	aData(idx).Add "名称","メンタルヘルス・マネジメント検定（Ⅱ種）"
	aData(idx).Add "名称TITLE","ﾒﾝﾀﾙﾍﾙｽ・ﾏﾈｼﾞﾒﾝﾄ検定（Ⅱ種）"
	aData(idx).Add "概要","【総務・人事・労務・事務の方にお薦めの資格】" & vbCrLf & vbCrLf & "企業内での適切なメンタルヘルス対策を講じる総務・人事労務管理に携わる方にお勧めの検定"
	aData(idx).Add "詳細","心の病を抱える労働者の増加が社会問題化しており、働く人の「心の健康管理」に関心が高まっています。厚生労働省では職場における適切かつ有効なメンタルヘルス対策の実施を推進しています。本検定試験は、厚生労働省の「労働者の心の健康の保持増進のための指針」に基づいて構築されています。"
	aData(idx).Add "費用","Ⅰ種　10,500円" & vbCrLf & "Ⅱ種　　6,300円" & vbCrlf & "Ⅲ種　　4,200円"
	aData(idx).Add "合格率","08年度" & vbCrLf & "Ⅰ種　11．1％" & vbCrLf & "Ⅱ種　70．7％" & vbCrLf & "Ⅲ種　87．4％"
	aData(idx).Add "団体名","大阪商工会議所"
	aData(idx).Add "教育機関","産業能率大学総合研究所"
	aData(idx).Add "講座内容","★通信教育" & vbCrLf & "（ﾃｷｽﾄ・添削有り）"
	aData(idx).Add "価格","受講料" & vbCrLf & "Ⅱ種　12,600円"
	aData(idx).Add "特典","特別受講料" & vbCrLf & "Ⅱ種　9,450円" & vbCrLf & "※支払方法：代金引き換えによる受講となります。"
	aData(idx).Add "クーポン",""
	aData(idx).Add "対象業種","総務・人事・労務・事務の方にお薦め"
	aData(idx).Add "価格2","Ⅱ種　12,600円"
	aData(idx).Add "概要2","企業内での適切なメンタルヘルス対策を講じる総務・人事労務管理に携わる方にお勧めの検定"
	aData(idx).Add "特典2","特別受講料<br>Ⅱ種　9,450円"

	'</ﾒﾝﾀﾙﾍﾙｽ・ﾏﾈｼﾞﾒﾝﾄ検定Ⅱ種>

	'<ﾒﾝﾀﾙﾍﾙｽ・ﾏﾈｼﾞﾒﾝﾄ検定Ⅲ種>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","license"
	aData(idx).Add "ID","0503"
	aData(idx).Add "優先",""
	aData(idx).Add "種別","ビジネス(資格)系 ／ 通信教育"
	aData(idx).Add "名称","メンタルヘルス" & vbCrLf & "マネジメント検定（Ⅲ種）"
	aData(idx).Add "名称TITLE","ﾒﾝﾀﾙﾍﾙｽ・ﾏﾈｼﾞﾒﾝﾄ検定（Ⅲ種）"
	aData(idx).Add "概要","【総務・人事・労務・事務の方にお薦めの資格】" & vbCrLf & vbCrLf & "企業内での適切なメンタルヘルス対策を講じる総務・人事労務管理に携わる方にお勧めの検定"
	aData(idx).Add "詳細","心の病を抱える労働者の増加が社会問題化しており、働く人の「心の健康管理」に関心が高まっています。厚生労働省では職場における適切かつ有効なメンタルヘルス対策の実施を推進しています。本検定試験は、厚生労働省の「労働者の心の健康の保持増進のための指針」に基づいて構築されています。"
	aData(idx).Add "費用","Ⅰ種　10,500円" & vbCrLf & "Ⅱ種　　6,300円" & vbCrlf & "Ⅲ種　　4,200円"
	aData(idx).Add "合格率","08年度" & vbCrLf & "Ⅰ種　11．1％" & vbCrLf & "Ⅱ種　70．7％" & vbCrLf & "Ⅲ種　87．4％"
	aData(idx).Add "団体名","大阪商工会議所"
	aData(idx).Add "教育機関","産業能率大学総合研究所"
	aData(idx).Add "講座内容","★通信教育" & vbCrLf & "（ﾃｷｽﾄ・添削有り）"
	aData(idx).Add "価格","受講料" & vbCrLf & "Ⅲ種　11,970円"
	aData(idx).Add "特典","特別受講料" & vbCrLf & "Ⅲ種　8,820円" & vbCrLf & "※支払方法：代金引き換えによる受講となります。"
	aData(idx).Add "クーポン",""
	aData(idx).Add "対象業種","総務・人事・労務・事務の方にお薦め"
	aData(idx).Add "価格2","Ⅲ種　11,970円"
	aData(idx).Add "概要2","企業内での適切なメンタルヘルス対策を講じる総務・人事労務管理に携わる方にお勧めの検定"
	aData(idx).Add "特典2","特別受講料<br>Ⅲ種　8,820円"

	'</ﾒﾝﾀﾙﾍﾙｽ・ﾏﾈｼﾞﾒﾝﾄ検定Ⅲ種>

	'<日商簿記検定2級>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","license"
	aData(idx).Add "ID","0602"
	aData(idx).Add "優先",""
	aData(idx).Add "種別","ビジネス(資格)系 ／ 通信教育"
	aData(idx).Add "名称","日商簿記検定試験（2級）"
	aData(idx).Add "名称TITLE","日商簿記検定試験（2級）"
	aData(idx).Add "概要","【経理はもちろんのこと多岐業務で活かせる資格】" & vbCrLf & vbCrLf & "会社経営の数字を理解するうえで必須のスキル"
	aData(idx).Add "詳細","企業規模や業種・業態を問わず、経営活動を記録・計算・整理し経営成績と財政状態を明らかにするだけでなく、取引先の経営状態をも把握できる技能であることから経理担当者だけでなく幅広いビジネススキルとして役立つ。他資格との組み合わせによりキャリアアップを目指すことが可能。"
	aData(idx).Add "費用","受験料は、2級　4,500円" & vbCrLf & "3級　2,500円"
	aData(idx).Add "合格率","2級" & vbCrlf & "Ｈ20.11実施　29.6％" & vbCrLf & "Ｈ20.06実施　31.3％" & vbCrLf & vbCrLf & "3級" & vbCrLf & "Ｈ20.11実施　40.2％" & vbCrlf & "Ｈ20.06実施　29.5％"
	aData(idx).Add "団体名","日本商工会議所"
	aData(idx).Add "教育機関","産業能率大学総合研究所"
	aData(idx).Add "講座内容","★通信教育" & vbCrlf & "（ﾃｷｽﾄ・添削有り）"
	aData(idx).Add "価格","受講料" & vbCrLf & "2級22,050円"
	aData(idx).Add "特典","約25％OFF" & vbCrLf & "特別受講料" & vbCrLf & "2級16,800円" & vbCrLf & "※試験の実施時期に合わせて、受験申請方法などの情報をお届け！" & vbCrLf & "※支払方法：代金引き換えによる受講となります。"
	aData(idx).Add "クーポン",""
	aData(idx).Add "対象業種","経理はもちろんのこと多岐業務で活かせる"
	aData(idx).Add "価格2","2級22,050円"
	aData(idx).Add "概要2","会社経営の数字を理解するうえで必須のスキル"
	aData(idx).Add "特典2","特別受講料<br>2級16,800円"

	'</日商簿記検定2級>

	'<日商簿記検定3級>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","license"
	aData(idx).Add "ID","0603"
	aData(idx).Add "優先",""
	aData(idx).Add "種別","ビジネス(資格)系 ／ 通信教育"
	aData(idx).Add "名称","日商簿記検定試験（3級）"
	aData(idx).Add "名称TITLE","日商簿記検定試験（3級）"
	aData(idx).Add "概要","【経理はもちろんのこと多岐業務で活かせる資格】" & vbCrLf & vbCrLf & "会社経営の数字を理解するうえで必須のスキル"
	aData(idx).Add "詳細","企業規模や業種・業態を問わず、経営活動を記録・計算・整理し経営成績と財政状態を明らかにするだけでなく、取引先の経営状態をも把握できる技能であることから経理担当者だけでなく幅広いビジネススキルとして役立つ。他資格との組み合わせによりキャリアアップを目指すことが可能。"
	aData(idx).Add "費用","受験料は、2級　4,500円" & vbCrLf & "3級　2,500円"
	aData(idx).Add "合格率","2級" & vbCrlf & "Ｈ20.11実施　29.6％" & vbCrLf & "Ｈ20.06実施　31.3％" & vbCrLf & vbCrLf & "3級" & vbCrLf & "Ｈ20.11実施　40.2％" & vbCrlf & "Ｈ20.06実施　29.5％"
	aData(idx).Add "団体名","日本商工会議所"
	aData(idx).Add "教育機関","産業能率大学総合研究所"
	aData(idx).Add "講座内容","★通信教育" & vbCrlf & "（ﾃｷｽﾄ・添削有り）"
	aData(idx).Add "価格","受講料" & vbCrLf & "3級19,950円　（一般<del>19,950円</del>）"
	aData(idx).Add "特典","約25％OFF" & vbCrLf & "特別受講料" & vbCrLf & "3級14,700円" & vbCrLf & "※試験の実施時期に合わせて、受験申請方法などの情報をお届け！" & vbCrLf & "※支払方法：代金引き換えによる受講となります。"
	aData(idx).Add "クーポン",""
	aData(idx).Add "対象業種","経理はもちろんのこと多岐業務で活かせる"
	aData(idx).Add "価格2","3級19,950円"
	aData(idx).Add "概要2","会社経営の数字を理解するうえで必須のスキル"
	aData(idx).Add "特典2","特別受講料<br>3級14,700円"

	'</日商簿記検定3級>

	'<ビジネス・キャリア検定試験2級>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","license"
	aData(idx).Add "ID","0102"
	aData(idx).Add "優先","1"
	aData(idx).Add "種別","ビジネス(資格)系 ／ 通信教育"
	aData(idx).Add "名称","キャリア検定試験（2級）"
	aData(idx).Add "名称TITLE","キャリア検定試験（2級）"
	aData(idx).Add "概要","【多岐業種対応の資格】" & vbCrLf & vbCrLf & "事務系職務を広く網羅した唯一の公的資格試験" &vbCrLf & "職務別に必要なスキル体系を構築し、企業実務に即した専門的知識・能力を客観的に評価した資格" & vbCrLf & vbCrLf & "レベルのイメージ" & vbCrLf & "職務に関連する幅広い総合的な専門知識を基に、グループやチームの中心メンバーとして、創意工夫を凝らし、自主的な判断・改善・提案を行いながら業務を遂行することができる。（例えば、課長、マネージャー等を目指す人、又はシニア・スタッフ）" & vbCrLf & vbCrLf & "◎分野は幅広い！" & vbCrLf & "人事人材開発・労務管理・経理・財務管理・経営情報システム、営業マーケティング" & vbCrLf & "など。"
	aData(idx).Add "詳細","ビジネスキャリア検定とは&nbsp;－&nbsp;人材力を高め、企業力を高めるビジネスキャリア検定" & vbCrLf & vbCrLf & "1．国の定める基準に準拠した試験・ビジネスキャリア検定" & vbCrLf & "国(厚生労働省)が、ビジネスパーソンの職務(セクション)別に必要なスキル体系(ガイドライン)を構築し、そのスキル体系(ガイドライン)を基準に公的資格としての検定試験を実施。" & vbCrLf & "" & vbCrLf & "2．職務を幅広くカバーした唯一の検定試験・ビジネスキャリア検定" & vbCrLf & "国(厚生労働省)と学識経験者グループが、ビジネスパーソンの職務(セクション)別に必要なスキル要素を抽出し、それらスキル要素の関連性･重要性などから体系を構築しているので職務全体の業務遂行のために必要なスキルが全網羅されている検定試験。" & vbCrLf & "" & vbCrLf & "3．実務能力の評価を重視した試験・ビジネスキャリア検定" & vbCrLf & "各職務に必要な知識修得をはじめ、試験は、1級、2級及び3級のレベルに体系化され、実務に即した専門的知識･能力を客観的に評価し、企業では社員の実務能力の客観的な評価や人材開発等に、個人にとって、キャリアアップなどに幅広く活用できる試験。" & vbCrLf & "" & vbCrLf & "4．学習しやすい体制・ビジネスキャリア検定" & vbCrLf & "厚生労働省から試験基準･ガイドラインに準拠した標準テキストを試験実施機関である中央職業能力開発協会が発行し、個人の自学習教材や通信、通学用講座の教材、企業内研修での教材として活用でき、中央職業能力開発協会が認定した試験対応講座を活用して学習することができる"
	aData(idx).Add "費用","※試験分野：" & vbCrLf & "人事人材開発（1.2.3級）、労務管理（1.2.3級）、経理（1.2.3級）、財務管理（1.2.3級）、経営情報システム（1.2.3級）、営業マーケティング（1.2.3級）" & vbCrLf & vbCrLf & "※費用：" & vbCrlf & "1級は7,850円、2級は5,250円、3級は4,200円"
	aData(idx).Add "合格率","人事人材開発（1.2.3級）→各20％.35％.43％" & vbCrLf & "労務管理（2.3級）→各24％.45％" & vbCrLf & "経理1級. 経理2級（財務会計）. 経理3級（原価計算）→各20％.39％.53％" & vbCrLf & "財務管理（2.3級）→各18％.68％" & vbCrLf & "経営情報システム（1級）→18％" & vbCrLf & "営業（1.2.3級）→20％.44％.41％" & vbCrLf & "その他一部非公開" & vbCrLf & "（08年前期）"
	aData(idx).Add "団体名","中央職業能力開発協会"
	aData(idx).Add "教育機関","【通信講座】" & vbCrLf & "【通学スクール】" & vbCrLf & "ＮＭＲビジネスキャリア学院"
	aData(idx).Add "講座内容","★通信教育（認定講座）" & vbCrLf & "※合格コース" & vbCrLf & "※自己啓発コース"
	aData(idx).Add "価格","★受講料" & vbCrLf & "分野別・コース別により異なる（22,000円～）" & vbCrLf & vbCrLf & "☆営業・マーケティング分野" & vbCrLf & "　・マーケティング2級　合格コース 38,000円　自己啓発コース 22,000円" & vbCrLf & "　・営業2級　合格コース 38,000円　自己啓発コース 22,000円" & vbCrLf & vbCrLf & "☆人事・人材開発・労務管理分野" & vbCrLf & "　・人事・人材開発2級　合格コース 38,000円　自己啓発コース 22,000円" & vbCrLf & "　・労務管理2級　合格コース 38,000円　自己啓発コース 22,000円" & vbCrLf & vbCrLf & "☆企業法務・総務分野" & vbCrLf & "　・企業法務2級(取引法務)　合格コース 38,000円　自己啓発コース 22,000円" & vbCrLf & "　・企業法務2級(組織法務)　合格コース 38,000円　自己啓発コース 22,000円" & vbCrLf & "　・総務2級　合格コース 38,000円　自己啓発コース 22,000円" & vbCrLf & vbCrLf & "☆経理・財務管理分野" & vbCrLf & "　・経理2級(財務会計)　合格コース 38,000円　自己啓発コース 22,000円" & vbCrLf & "　・財務管理2級(財務管理・管理会計)　合格コース 38,000円　自己啓発コース 22,000円" & vbCrLf & vbCrLf & "☆経営戦略分野" & vbCrLf & "　・経営戦略2級　合格コース 38,000円　自己啓発コース 22,000円" & vbCrLf & vbCrLf & "☆経営情報システム分野" & vbCrLf & "　・経営情報システム2級(情報化企画)　合格コース 38,000円　自己啓発コース 22,000円" & vbCrLf & "　・経営情報システム2級(情報化活用)　合格コース 38,000円　自己啓発コース 22,000円" & vbCrLf & vbCrLf & "☆ロジスティクス分野" & vbCrLf & "　・ロジスティクス管理2級　合格コース 38,000円　自己啓発コース 22,000円" & vbCrLf & "　・ロジスティクス・オペレーション2級　合格コース 38,000円　自己啓発コース 22,000円" & vbCrLf & vbCrLf & "☆生産管理分野" & vbCrLf & "　・生産管理プランニング2級(生産システム・生産管理)　合格コース 38,000円　自己啓発コース 22,000円" & vbCrLf & "　・生産管理プランニング2級(製品企画・設計管理)　合格コース 38,000円　自己啓発コース 22,000円" & vbCrLf & "　・生産管理オペレーション2級(購買・物流・在庫管理)　合格コース 38,000円　自己啓発コース 22,000円" & vbCrLf & "　・生産管理オペレーション2級(作業・工程・設備管理)　合格コース 38,000円　自己啓発コース 22,000円"
	aData(idx).Add "特典",""
	aData(idx).Add "クーポン",""
	aData(idx).Add "クーポン注意点",""
	aData(idx).Add "対象業種","多岐業種対応"
	aData(idx).Add "価格2","コース別により異なる<br>（22,000円～）"
	aData(idx).Add "概要2","事務系職務を広く網羅した唯一の公的資格試験"
	aData(idx).Add "特典2",""

	'</ビジネス・キャリア検定試験2級>

	'<ビジネス・キャリア検定試験3級>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","license"
	aData(idx).Add "ID","0103"
	aData(idx).Add "優先","1"
	aData(idx).Add "種別","ビジネス(資格)系 ／ 通信教育"
	aData(idx).Add "名称","キャリア検定試験（3級）"
	aData(idx).Add "名称TITLE","キャリア検定試験（3級）"
	aData(idx).Add "概要","【多岐業種対応の資格】" & vbCrLf & vbCrLf & "事務系職務を広く網羅した唯一の公的資格試験" &vbCrLf & "職務別に必要なスキル体系を構築し、企業実務に即した専門的知識・能力を客観的に評価した資格" & vbCrLf & vbCrLf & "レベルのイメージ" & vbCrLf & "職務全般に関する幅広い専門知識を基に、担当者として上司の指示・助言を踏まえ、自ら問題意識を持ち定型的業務を確実に遂行することができる。（例えば、係長、リーダー等を目指す人、又は担当業務を的確に遂行できることを目指す人）" & vbCrLf & vbCrLf & "◎分野は幅広い！" & vbCrLf & "人事人材開発・労務管理・経理・財務管理・経営情報システム、営業マーケティング" & vbCrLf & "など。"
	aData(idx).Add "詳細","ビジネスキャリア検定とは&nbsp;－&nbsp;人材力を高め、企業力を高めるビジネスキャリア検定" & vbCrLf & vbCrLf & "1．国の定める基準に準拠した試験・ビジネスキャリア検定" & vbCrLf & "国(厚生労働省)が、ビジネスパーソンの職務(セクション)別に必要なスキル体系(ガイドライン)を構築し、そのスキル体系(ガイドライン)を基準に公的資格としての検定試験を実施。" & vbCrLf & "" & vbCrLf & "2．職務を幅広くカバーした唯一の検定試験・ビジネスキャリア検定" & vbCrLf & "国(厚生労働省)と学識経験者グループが、ビジネスパーソンの職務(セクション)別に必要なスキル要素を抽出し、それらスキル要素の関連性･重要性などから体系を構築しているので職務全体の業務遂行のために必要なスキルが全網羅されている検定試験。" & vbCrLf & "" & vbCrLf & "3．実務能力の評価を重視した試験・ビジネスキャリア検定" & vbCrLf & "各職務に必要な知識修得をはじめ、試験は、1級、2級及び3級のレベルに体系化され、実務に即した専門的知識･能力を客観的に評価し、企業では社員の実務能力の客観的な評価や人材開発等に、個人にとって、キャリアアップなどに幅広く活用できる試験。" & vbCrLf & "" & vbCrLf & "4．学習しやすい体制・ビジネスキャリア検定" & vbCrLf & "厚生労働省から試験基準･ガイドラインに準拠した標準テキストを試験実施機関である中央職業能力開発協会が発行し、個人の自学習教材や通信、通学用講座の教材、企業内研修での教材として活用でき、中央職業能力開発協会が認定した試験対応講座を活用して学習することができる"
	aData(idx).Add "費用","※試験分野：" & vbCrLf & "人事人材開発（1.2.3級）、労務管理（1.2.3級）、経理（1.2.3級）、財務管理（1.2.3級）、経営情報システム（1.2.3級）、営業マーケティング（1.2.3級）" & vbCrLf & vbCrLf & "※費用：" & vbCrlf & "1級は7,850円、2級は5,250円、3級は4,200円"
	aData(idx).Add "合格率","人事人材開発（1.2.3級）→各20％.35％.43％" & vbCrLf & "労務管理（2.3級）→各24％.45％" & vbCrLf & "経理1級. 経理2級（財務会計）. 経理3級（原価計算）→各20％.39％.53％" & vbCrLf & "財務管理（2.3級）→各18％.68％" & vbCrLf & "経営情報システム（1級）→18％" & vbCrLf & "営業（1.2.3級）→20％.44％.41％" & vbCrLf & "その他一部非公開" & vbCrLf & "（08年前期）"
	aData(idx).Add "団体名","中央職業能力開発協会"
	aData(idx).Add "教育機関","【通信講座】" & vbCrLf & "【通学スクール】" & vbCrLf & "ＮＭＲビジネスキャリア学院"
	aData(idx).Add "講座内容","★通信教育（認定講座）" & vbCrLf & "※合格コース" & vbCrLf & "※自己啓発コース"
	aData(idx).Add "価格","★受講料" & vbCrLf & "分野別・コース別により異なる（19,500円～）" & vbCrLf & vbCrLf & "☆営業・マーケティング分野" & vbCrLf & "　・マーケティング3級　合格コース 33,000円　自己啓発コース 19,500円" & vbCrLf & "　・営業3級　合格コース 33,000円　自己啓発コース 19,500円" & vbCrLf & vbCrLf & "☆人事・人材開発・労務管理分野" & vbCrLf & "　・人事・人材開発3級　合格コース 33,000円　自己啓発コース 19,500円" & vbCrLf & "　・労務管理3級　合格コース 33,000円　自己啓発コース 19,500円" & vbCrLf & vbCrLf & "☆企業法務・総務分野" & vbCrLf & "　・企業法務3級　合格コース 33,000円　自己啓発コース 19,500円" & vbCrLf & "　・総務3級　合格コース 33,000円　自己啓発コース 19,500円" & vbCrLf & vbCrLf & "☆経理・財務管理分野" & vbCrLf & "　・経理3級(簿記・財務諸表)　合格コース 33,000円　自己啓発コース 19,500円" & vbCrLf & "　・経理3級(原価計算)　合格コース 33,000円　自己啓発コース 19,500円" & vbCrLf & "　・財務管理3級　合格コース 33,000円　自己啓発コース 19,500円" & vbCrLf & vbCrLf & "☆経営戦略分野" & vbCrLf & "　・経営戦略3級　合格コース 33,000円　自己啓発コース 19,500円" & vbCrLf & vbCrLf & "☆経営情報システム分野" & vbCrLf & "　・経営情報システム3級　合格コース 33,000円　自己啓発コース 19,500円" & vbCrLf & vbCrLf & "☆ロジスティクス分野" & vbCrLf & "　・ロジスティクス管理3級　合格コース 33,000円　自己啓発コース 19,500円" & vbCrLf & "　・ロジスティクス・オペレーション3級　合格コース 33,000円　自己啓発コース 19,500円" & vbCrLf & vbCrLf & "☆生産管理分野" & vbCrLf & "　・生産管理プランニング3級　合格コース 33,000円　自己啓発コース 19,500円" & vbCrLf & "　・生産管理オペレーション3級　合格コース 33,000円　自己啓発コース 19,500円"
	aData(idx).Add "特典",""
	aData(idx).Add "クーポン",""
	aData(idx).Add "クーポン注意点",""
	aData(idx).Add "対象業種","多岐業種対応"
	aData(idx).Add "価格2","コース別により異なる<br>（19,500円～）"
	aData(idx).Add "概要2","事務系職務を広く網羅した唯一の公的資格試験"
	aData(idx).Add "特典2",""

	'</ビジネス・キャリア検定試験3級>

	'<衛生管理者1種>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","license"
	aData(idx).Add "ID","0201"
	aData(idx).Add "優先",""
	aData(idx).Add "種別","ビジネス(資格)系 ／ 通信教育"
	aData(idx).Add "名称","衛生管理者（1種）"
	aData(idx).Add "名称TITLE","衛生管理者（1種）"
	aData(idx).Add "概要","【全業種対応の資格】" & vbCrLf & "【総務・人事・事務の方にお薦め】" & vbCrLf & vbCrLf & "『国家資格』(厚生労働省認定）として、全業種に選任が義務づけられている資格"
	aData(idx).Add "詳細","常時50人以上の労働者を使用する事業場では、衛生管理者免許を有する者のうちから労働者数に応じて一定数以上の衛生管理者を選任し、安全衛生業務のうち、衛生に関わる技術的な事項を管理させることが必要になります。" & vbCrLf & "第1種は全ての事業所において管理者になれます。第2種は、有害業務と関連の薄い情報通信業や金融業などの一定の業種の事業場においてのみ管理者になれます。おもな職務は、労働者の健康障害を防止するための作業環境管理・作業管理・健康管理・労働衛生教育の実施・健康保持増進措置などです。"
	aData(idx).Add "費用","受験料は8,300円。" & vbCrLf & "合格後、登録手数料印紙代1,500円他がかかる。"
	aData(idx).Add "合格率","第1種：54.7％" & vbCrLf & "第2種：65.6％" & vbCrLf & "（07年度）"
	aData(idx).Add "団体名","（財）安全衛生技術試験協会"
	aData(idx).Add "教育機関","産業能率大学総合研究所"
	aData(idx).Add "講座内容","★通信教育" & vbCrLf & "（ﾃｷｽﾄ・添削有り）" & vbCrLf & "※予想問題集・用語解説集付"
	aData(idx).Add "価格","★受講料" & vbCrLf & "1種　25,200円" & vbCrLf & "※本試験に対応した実践的リポート問題で確実に合格支援！" & vbCrLf & "※支払方法：代金引き換えによる受講となります。"
	aData(idx).Add "特典",""
	aData(idx).Add "クーポン",""
	aData(idx).Add "対象業種","全業種対応"
	aData(idx).Add "価格2","1種　25,200円"
	aData(idx).Add "概要2","『国家資格』(厚生労働省認定）として、全業種に選任が義務づけられている資格"
	aData(idx).Add "特典2",""

	'</衛生管理者1種>

	'<衛生管理者2種>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","license"
	aData(idx).Add "ID","0202"
	aData(idx).Add "優先",""
	aData(idx).Add "種別","ビジネス(資格)系 ／ 通信教育"
	aData(idx).Add "名称","衛生管理者（2種）"
	aData(idx).Add "名称TITLE","衛生管理者（2種）"
	aData(idx).Add "概要","【全業種対応の資格】" & vbCrLf & "【総務・人事・事務の方にお薦め】" & vbCrLf & vbCrLf & "『国家資格』(厚生労働省認定）として、全業種に選任が義務付けられている資格"
	aData(idx).Add "詳細","常時50人以上の労働者を使用する事業場では、衛生管理者免許を有する者のうちから労働者数に応じて一定数以上の衛生管理者を選任し、安全衛生業務のうち、衛生に関わる技術的な事項を管理させることが必要になります。" & vbCrLf & "第1種は全ての事業所において管理者になれます。第2種は、有害業務と関連の薄い情報通信業や金融業などの一定の業種の事業場においてのみ管理者になれます。おもな職務は、労働者の健康障害を防止するための作業環境管理・作業管理・健康管理・労働衛生教育の実施・健康保持増進措置などです。"
	aData(idx).Add "費用","受験料は8,300円。" & vbCrLf & "合格後、登録手数料印紙代1,500円他がかかる。"
	aData(idx).Add "合格率","第1種：54.7％" & vbCrLf & "第2種：65.6％" & vbCrLf & "（07年度）"
	aData(idx).Add "団体名","（財）安全衛生技術試験協会"
	aData(idx).Add "教育機関","産業能率大学総合研究所"
	aData(idx).Add "講座内容","★通信教育" & vbCrLf & "（ﾃｷｽﾄ・添削有り）" & vbCrLf & "※予想問題集・用語解説集付"
	aData(idx).Add "価格","★受講料" & vbCrLf & "2種　23,100円" & vbCrLf & "※本試験に対応した実践的リポート問題で確実に合格支援！" & vbCrLf & "※支払方法：代金引き換えによる受講となります。"
	aData(idx).Add "特典",""
	aData(idx).Add "クーポン",""
	aData(idx).Add "対象業種","全業種対応"
	aData(idx).Add "価格2","2種　23,100円"
	aData(idx).Add "概要2","『国家資格』(厚生労働省認定）として、全業種に選任が義務付けられている資格"
	aData(idx).Add "特典2",""

	'</衛生管理者2種>

	'<情報処理技術者試験　ＩＴパスポート>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","license"
	aData(idx).Add "ID","0301"
	aData(idx).Add "優先",""
	aData(idx).Add "種別","ビジネス(資格)系 ／ 通信教育"
	aData(idx).Add "名称","情報処理技術者試験"
	aData(idx).Add "名称TITLE","情報処理技術者試験"
	aData(idx).Add "概要","【全業種対応の資格】" & vbCrLf & vbCrLf & "『国家資格』（経済産業省認定）として、初級シスアド試験に代わるパソコンを使う全ての人が対象の資格"
	aData(idx).Add "詳細","職業人誰もが共通に備えておくべき情報技術に関する基礎的な知識を測る、情報処理技術者試験のレベル1の試験。" & vbCrLf & "日経ｿﾘｭｰｼｮﾝﾋﾞｼﾞﾈｽが実施したアンケート調査（2009年版「いる資格、いらない資格」）では、営業職に取らせたい資格の上位10位中6つが情報処理技術者試験が占めていた。ITパスポート資格は第2位。" & vbCrLf & "第1位は、情報処理技術者試験　基本情報技術者試験（レベル2）"
	aData(idx).Add "費用","受験料は　5,100円。"
	aData(idx).Add "合格率","31．0％（初級ｼｽｱﾄﾞ試験）"
	aData(idx).Add "団体名","（独）情報処理推進機構"
	aData(idx).Add "教育機関","産業能率大学総合研究所"
	aData(idx).Add "講座内容","★通信教育" & vbCrLf & "（テキスト・問題集・実力テスト3回付き）"
	aData(idx).Add "価格","★受講料" & vbCrLf & "18,900円" & vbCrLf & "※新試験制度のエントリレベルに該当する試験合格を目指すコース！" & vbCrLf & "※支払方法：代金引き換えによる受講となります。"
	aData(idx).Add "特典",""
	aData(idx).Add "クーポン",""
	aData(idx).Add "対象業種","全業種対応"
	aData(idx).Add "価格2","18,900円"
	aData(idx).Add "概要2","『国家資格』（経済産業省認定）として、初級シスアド試験に代わるパソコンを使う全ての人が対象の資格"
	aData(idx).Add "特典2",""

	'</情報処理技術者試験　ＩＴパスポート>

	'<登録販売者>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","license"
	aData(idx).Add "ID","0601"
	aData(idx).Add "優先",""
	aData(idx).Add "種別","ビジネス(資格)系 ／ 通信教育"
	aData(idx).Add "名称","登録販売者"
	aData(idx).Add "名称TITLE","登録販売者"
	aData(idx).Add "概要","【薬局で市販薬販売のお仕事をご希望の方にお薦めの資格】" & vbCrLf & vbCrLf & "最新！2009年6月改正薬事法の新資格！"
	aData(idx).Add "詳細","2009年度より改正薬事法が施行されています。一般用医薬品は副作用リスクに応じて3分類され、リスクの低い第二類・第三類は「登録販売者」が販売できるようになりました。" & vbCrLf & "登録販売者は都道府県が実施する登録販売者試験に合格することにより取得できます。" & vbCrLf & "ドラッグストア・コンビニ業界では、この「登録販売者」の確保と育成が重要な人材開発課題になっていくことが予想されています！"
	aData(idx).Add "費用",""
	aData(idx).Add "合格率","2008年度　第1回試験合格率（薬局新聞2009年4月1日付より）" & vbCrLf & vbCrLf & "東京　82．3 (％)" & vbCrLf & vbCrLf & "埼玉　77．0 (％)" & vbCrLf & "千葉　80．0 (％)" & vbCrLf & "神奈川　84．5 (％)" & vbCrLf & vbCrLf & "北海道　54．8 (％)" & vbCrLf & "青森　53．1 (％)" & vbCrLf & "岩手　43．0 (％)" & vbCrLf & "宮城　53．6 (％)" & vbCrLf & "秋田　52．9 (％)" & vbCrLf & "山形　47．5 (％)" & vbCrLf & "福島　52．2 (％)" & vbCrLf & vbCrLf & "新潟　75．4 (％)" & vbCrLf & "山梨　66．6 (％)" & vbCrLf & "長野　75．5 (％)" & vbCrLf & "茨城　73．8 (％)" & vbCrLf & "栃木　71．1 (％)" & vbCrLf & "群馬　77．6 (％)" & vbCrLf & vbCrLf & "福岡　63．2 (％)" & vbCrLf & "佐賀　55．7 (％)" & vbCrLf & "熊本　62．9 (％)" & vbCrLf & "大分　54．5 (％)" & vbCrLf & "宮崎　63．9 (％)" & vbCrLf & "鹿児島　56．0 (％)" & vbCrLf & "沖縄　47．8 (％)"
	aData(idx).Add "団体名","各都道府県"
	aData(idx).Add "教育機関","産業能率大学総合研究所"
	aData(idx).Add "講座内容","★通信教育" & vbCrLf & "（ﾃｷｽﾄ・添削あり）"
	aData(idx).Add "価格","受講料" & vbCrLf & "22,050円"
	aData(idx).Add "特典","特別受講料" & vbCrLf & "16,800円"
	aData(idx).Add "クーポン",""
	aData(idx).Add "対象業種","薬局で市販薬販売のお仕事をご希望の方"
	aData(idx).Add "価格2","22,050円"
	aData(idx).Add "概要2","2009年度より改正薬事法が施行されています。一般用医薬品は副作用リスクに応じて3分類され、リスクの低い第二類・第三類は「登録販売者」が販売できるようになりました。"
	aData(idx).Add "特典2","特別受講料<br>16,800円"

	'</登録販売者>

	'<テンプレート>
	'idx = idx + 1
	'Set aData(idx) = Server.CreateObject("scripting.dictionary")
	'aData(idx).Add "Category","skillup"
	'aData(idx).Add "ID",""
	'aData(idx).Add "優先",""
	'aData(idx).Add "種別",""
	'aData(idx).Add "名称",""
	'aData(idx).Add "教育機関",""
	'aData(idx).Add "講座内容",""
	'aData(idx).Add "価格",""
	'aData(idx).Add "特典",""
	'aData(idx).Add "クーポン",""
	'aData(idx).Add "クーポン注意点",""
	'aData(idx).Add "対象業種",""
	'aData(idx).Add "価格2",""
	'aData(idx).Add "概要2",""
	'aData(idx).Add "特典2",""

	'</テンプレート>

	'<語学スキル>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","skillup"
	aData(idx).Add "ID","0101"
	aData(idx).Add "優先","1"
	aData(idx).Add "種別","ビジネス(資格)系 ／ 通学"
	aData(idx).Add "名称","英会話（マンツーマン）"
	aData(idx).Add "名称TITLE","英会話（マンツーマン）"
	aData(idx).Add "教育機関","Gabaマンツーマン英会話"
	aData(idx).Add "講座内容","【通学スクール】"
	aData(idx).Add "価格",""
	aData(idx).Add "特典","1．入会金 ￥10,500 （通常 ￥31,500）" & vbCrLf & "2．レッスン料金から ￥21,000OFF（対象コース：63回以上）" & vbCrLf & "3．Lesson Anywhere 無料進呈（通常 ￥4,200）" & vbCrLf & "　　※どのスクールでもレッスンを受講できるオプションです。"
	aData(idx).Add "クーポン",""
	aData(idx).Add "クーポン注意点",""
	aData(idx).Add "対象業種","全業種対応"
	aData(idx).Add "価格2",""
	aData(idx).Add "概要2",""
	aData(idx).Add "特典2","<div style=""text-align:left; font-size:11px;"">1．入会金 ￥10,500<br>2．レッスン料金<br>￥21,000OFF<br>3．Lesson Anywhere<br>無料進呈"

	'</語学スキル>

	'<パソコンスキルアップ>
	'idx = idx + 1
	'Set aData(idx) = Server.CreateObject("scripting.dictionary")
	'aData(idx).Add "Category","skillup"
	'aData(idx).Add "ID","0301"
	'aData(idx).Add "優先",""
	'aData(idx).Add "種別","ビジネス(資格)系 ／ 通学"
	'aData(idx).Add "名称","パソコンスキルアップ"
	'aData(idx).Add "教育機関","アビバ"
	'aData(idx).Add "講座内容","【通学スクール】"
	'aData(idx).Add "価格",""
	'aData(idx).Add "特典","全国160校どこでも特典利用可能" & vbCrLf & "入会金2万円以上OFF" & vbCrLf & "受講料特別割引5％OFF！"
	'aData(idx).Add "クーポン","1"
	'aData(idx).Add "クーポン注意点","・クーポンをご利用の際は、必ず身分証明書をご持参下さい。" & vbCrLf & "・スクールでカウンセリングを行い、コース設定をした上で特典が適用されます。"
	'aData(idx).Add "対象業種",""
	'aData(idx).Add "価格2",""
	'aData(idx).Add "概要2",""
	'aData(idx).Add "特典2",""

	'</パソコンスキルアップ>

	'<パソコンスキルアップ>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","skillup"
	aData(idx).Add "ID","0401"
	aData(idx).Add "優先",""
	aData(idx).Add "種別","ビジネス(資格)系 ／ 通学"
	aData(idx).Add "名称","パソコンスキルアップ"
	aData(idx).Add "名称TITLE","パソコンスキルアップ"
	aData(idx).Add "教育機関","産業能率大学総合研究所"
	aData(idx).Add "講座内容","【通信講座】"
	aData(idx).Add "価格","□総合コース" & vbCrLf & "　技ありシリーズ" & vbCrLf & "　　【Office技あり〈Office2003・2002・2000〉】　一般価格 16,800円 ※30％以上OFF" & vbCrLf & "　　※Word・Excel・PowerPointを使いこなすコースです！" & vbCrLf & "" & vbCrLf & "□Excel・Wordコース" & vbCrLf & "　見て・やって・簡単" & vbCrLf & "　　【Excel2002総合】　一般価格 19,950円" & vbCrLf & "　　【Word2002総合】　一般価格 19,950円" & vbCrLf & "　　【Excel・Word2002基礎】　一般価格 19,950円" & vbCrLf & "　　【Excel・Word2003応用】　一般価格 19,950円" & vbCrLf & "" & vbCrLf & "□PowerPointコース" & vbCrLf & "　【使えるプレゼン！PowerPoint】　一般価格 19,950円" & vbCrLf & "" & vbCrLf & "□インターネットコース" & vbCrLf & "　【インターネットセキュリティー”超”入門】　一般価格 13,650円" & vbCrLf & "　【インターネット活用】　一般価格 13,650円"
	aData(idx).Add "特典","□総合コース" & vbCrLf & "　技ありシリーズ" & vbCrLf & "　　【Office技あり〈Office2003・2002・2000〉】　特別価格 11,550円(通常：16,800円) ※30％以上OFF" & vbCrLf & "　　※Word・Excel・PowerＰｏｉｎｔを使いこなすコースです！" & vbCrLf & "" & vbCrLf & "□Excel・Wordコース" & vbCrLf & "　見て・やって・簡単" & vbCrLf & "　　【Excel2002総合】　特別価格 14,700円(通常：19,950円) ※26％OFF" & vbCrLf & "　　【Word2002総合】　特別価格 14,700円(通常：19,950円) ※26％OFF" & vbCrLf & "　　【Excel・Word2002基礎】　特別価格 14,700円(通常：19,950円) ※26％OFF" & vbCrLf & "　　【Excel・Word2003応用】　特別価格 14,700円(通常：19,950円) ※26％OFF" & vbCrLf & "" & vbCrLf & "□PowerPointコース" & vbCrLf & "　【使えるプレゼン！PowerPoint】　特別価格 14,700円(通常：19,950円) ※26％OFF" & vbCrLf & "" & vbCrLf & "□インターネットコース" & vbCrLf & "　【インターネットセキュリティー”超”入門】　特別価格 8,400円(通常：13,650円) ※38％OFF" & vbCrLf & "　【インターネット活用】　特別価格 8,400円(通常：13,650円) ※38％OFF"
	aData(idx).Add "クーポン",""
	aData(idx).Add "クーポン注意点",""
	aData(idx).Add "対象業種","全業種対応"
	aData(idx).Add "価格2","Office技ありコースなど<br>(通常：16,800円～)"
	aData(idx).Add "概要2",""
	aData(idx).Add "特典2","特別価格<br>最大38％OFF<br>など"

	'</パソコンスキルアップ>

	'<語学スキル>
	idx = idx + 1
	Set aData(idx) = Server.CreateObject("scripting.dictionary")
	aData(idx).Add "Category","skillup"
	aData(idx).Add "ID","0201"
	aData(idx).Add "優先",""
	aData(idx).Add "種別","ビジネス(資格)系 ／ 通信教育"
	aData(idx).Add "名称","TOEIC実践トレーニング" & vbCrLf & "750ｸﾘｱ・650ｸﾘｱ・550ｸﾘｱ・450ｸﾘｱの各コース設定"
	aData(idx).Add "名称TITLE","TOEIC実践トレーニング"
	aData(idx).Add "教育機関","産業能率大学総合研究所"
	aData(idx).Add "講座内容","【通信講座】"
	aData(idx).Add "価格","750クリア → 31,500円" & vbCrLf & "650クリア → 23,100円" & vbCrLf & "550クリア → 22,050円" & vbCrLf & "450クリア → 21,000円" & vbCrLf & "※支払方法：代金引き換えによる受講となります。"
	aData(idx).Add "特典",""
	aData(idx).Add "クーポン",""
	aData(idx).Add "クーポン注意点",""
	aData(idx).Add "対象業種","全業種対応"
	aData(idx).Add "価格2","750クリア → 31,500円<br>650クリア → 23,100円<br>550クリア → 22,050円<br>450クリア → 21,000円"
	aData(idx).Add "概要2",""
	aData(idx).Add "特典2",""

	'</語学スキル>

	getWarmreceptionMaster = aData
End Function
%>

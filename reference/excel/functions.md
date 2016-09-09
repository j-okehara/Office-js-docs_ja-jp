﻿# ワークシート関数
JavaScript または REST から呼び出すことができる Microsoft Excel ワークシート関数のコンテナーとして使用されます。

## 戻り値の型
ワークシート関数は、FunctionResult オブジェクトを返します。FunctionResult オブジェクトには、2 つのプロパティがあります。

| プロパティ       | 型    |説明|メモ |
|:---------------|:--------|:----------|:-----|
|value|string|指定されたワークシート関数を適用した結果を返します。||
|エラー|string|指定されたワークシート関数を適用したときのエラー情報を返します。||


## 構文
```js
workbook.functions.functionMethod();
```

## 例
次のような、各種の工具の売上を追跡するサンプル表 (Excel) があるとします。この表のデータを使用して、ワークシート関数が機能するしくみについて説明します。

![サンプル](../../images/worksheetfunctionschainingResult.JPG)


次の例では、vlookup 関数を前述の表に適用し、11 月のレンチの売上を検索します。
```js
    Excel.run(function (ctx) {
        var range = ctx.workbook.worksheets.getItem("Sheet1").getRange("A1:D4");
        var unitSoldInNov = ctx.workbook.functions.vlookup("Wrench", range, 2, false);
        unitSoldInNov.load();
        return ctx.sync()
        .then(function () {
            console.log(unitSoldInNov.value);
        });
    })
    .catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
```




次の例では、vlookup 関数を使用して、最初に 11 月と 12 月のレンチの売上を別個に検出します。次に、sum 関数を適用して、その 2 か月の売上合計を出します。読み込む必要があるのは最終的な結果だけで、途中の結果は最後の数式を適用するときに計算され使用される点に注意してください。

```js
    Excel.run(function (ctx) {
        var range = ctx.workbook.worksheets.getItem("Sheet1").getRange("A1:D4");
        var sumOfTwoLookups = ctx.workbook.functions.sum(
            ctx.workbook.functions.vlookup("Wrench", range, 2, false), 
            ctx.workbook.functions.vlookup("Wrench", range, 3, false)
            );
        sumOfTwoLookups.load();
        return ctx.sync()
        .then(function () {
            console.log(sumOfTwoLookups.value);
        });
    })
    .catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });


```

## [サポートされているワークシート関数の一覧](#サポートされているワークシート関数の一覧)

| メソッド           | 戻り値の型    |説明|メモ |
|:---------------|:--------|:----------|:-----|
|[ABS 関数](https://support.office.com/en-us/article/ABS-function-3420200f-5628-4e8c-99da-c99d7c87713c)| FunctionResult |数値の絶対値を返します。|
|[ACCRINT 関数](https://support.office.com/en-us/article/ACCRINT-function-fe45d089-6722-4fb3-9379-e1f911d8dc74)| FunctionResult |定期的に利息が支払われる証券の未収利息額を返します。|
|[ACCRINTM 関数](https://support.office.com/en-us/article/ACCRINTM-function-f62f01f9-5754-4cc4-805b-0e70199328a7)| FunctionResult |満期日に利息が支払われる証券の未収利息額を返します。|
|[ACOS 関数](https://support.office.com/en-us/article/ACOS-function-cb73173f-d089-4582-afa1-76e5524b5d5b)| FunctionResult |数値の逆余弦 (アークコサイン) を返します。|
|[ACOSH 関数](https://support.office.com/en-us/article/ACOSH-function-e3992cc1-103f-4e72-9f04-624b9ef5ebfe)| FunctionResult |数値の逆双曲線余弦を返します。|
|[ACOT 関数](https://support.office.com/en-us/article/ACOT-function-dc7e5008-fe6b-402e-bdd6-2eea8383d905)| FunctionResult |数値の逆余接 (アークコタンジェント) を返します。|
|[ACOTH 関数](https://support.office.com/en-us/article/ACOTH-function-cc49480f-f684-4171-9fc5-73e4e852300f)| FunctionResult |数値の逆双曲線余接を返します。|
|[AMORDEGRC 関数](https://support.office.com/en-us/article/AMORDEGRC-function-a14d0ca1-64a4-42eb-9b3d-b0dededf9e51)| FunctionResult |減価償却係数を使用して、各会計期における減価償却費を返します。|
|[AMORLINC 関数](https://support.office.com/en-us/article/AMORLINC-function-7d417b45-f7f5-4dba-a0a5-3451a81079a8)| FunctionResult |各会計期における減価償却費を返します。|
|[AND 関数](https://support.office.com/en-us/article/AND-function-5f19b2e8-e1df-4408-897a-ce285a19e9d9)| FunctionResult |すべての引数が TRUE のときに TRUE を返します。|
|[ARABIC 関数](https://support.office.com/en-us/article/ARABIC-function-9a8da418-c17b-4ef9-a657-9370a30a674f)| FunctionResult |ローマ数字をアラビア数字に変換します。|
|[AREAS 関数](https://support.office.com/en-us/article/AREAS-function-8392ba32-7a41-43b3-96b0-3695d2ec6152)| FunctionResult |指定の範囲に含まれる領域の個数を返します。|
|[ASC 関数](https://support.office.com/en-us/article/ASC-function-0b6abf1c-c663-4004-a964-ebc00b723266)| FunctionResult |全角 (2 バイト) の英数カナ文字を半角 (1 バイト) の文字に変換します。|
|[ASIN 関数](https://support.office.com/en-us/article/ASIN-function-81fb95e5-6d6f-48c4-bc45-58f955c6d347)| FunctionResult |数値の逆正弦 (アークサイン) を返します。|
|[ASINH 関数](https://support.office.com/en-us/article/ASINH-function-4e00475a-067a-43cf-926a-765b0249717c)| FunctionResult |数値の逆双曲線正弦を返します。|
|[ATAN 関数](https://support.office.com/en-us/article/ATAN-function-50746fa8-630a-406b-81d0-4a2aed395543)| FunctionResult |数値の逆正接 (アークタンジェント) を返します。|
|[ATAN2 関数](https://support.office.com/en-us/article/ATAN2-function-c04592ab-b9e3-4908-b428-c96b3a565033)| FunctionResult |指定された x-y 座標の逆正接 (アークタンジェント) を返します。|
|[ATANH 関数](https://support.office.com/en-us/article/ATANH-function-3cd65768-0de7-4f1d-b312-d01c8c930d90)| FunctionResult |数値の逆双曲線正接を返します。|
|[AVEDEV 関数](https://support.office.com/en-us/article/AVEDEV-function-58fe8d65-2a84-4dc7-8052-f3f87b5c6639)| FunctionResult |データ全体の平均値に対するそれぞれのデータの絶対偏差の平均を返します。|
|[AVERAGE 関数](https://support.office.com/en-us/article/AVERAGE-function-047bac88-d466-426c-a32b-8f33eb960cf6)| FunctionResult |引数の平均値を返します。|
|[AVERAGEA 関数](https://support.office.com/en-us/article/AVERAGEA-function-f5f84098-d453-4f4c-bbba-3d2c66356091)| FunctionResult |数値、文字列、および論理値を含む引数の平均値を返します。|
|[AVERAGEIF 関数](https://support.office.com/en-us/article/AVERAGEIF-function-faec8e2e-0dec-4308-af69-f5576d8ac642)| FunctionResult |範囲内の検索条件に一致するすべてのセルの平均値 (算術平均) を返します。|
|[AVERAGEIFS 関数](https://support.office.com/en-us/article/AVERAGEIFS-function-48910c45-1fc0-4389-a028-f7c5c3001690)| FunctionResult |複数の検索条件に一致するすべてのセルの平均値 (算術平均) を返します。|
|[BAHTTEXT 関数](https://support.office.com/en-us/article/BAHTTEXT-function-5ba4d0b4-abd3-4325-8d22-7a92d59aab9c)| FunctionResult |バーツ (ß) 通貨書式を使用して、数値を文字列に変換します。|
|[BASE 関数](https://support.office.com/en-us/article/BASE-function-2ef61411-aee9-4f29-a811-1c42456c6342)| FunctionResult |数値を、指定された基数 (底) のテキスト表現に変換します。|
|[BESSELI 関数](https://support.office.com/en-us/article/BESSELI-function-8d33855c-9a8d-444b-98e0-852267b1c0df)| FunctionResult |修正ベッセル関数 In(x) を返します。|
|[BESSELJ 関数](https://support.office.com/en-us/article/BESSELJ-function-839cb181-48de-408b-9d80-bd02982d94f7)| FunctionResult |ベッセル関数 Jn(x) を返します。|
|[BESSELK 関数](https://support.office.com/en-us/article/BESSELK-function-606d11bc-06d3-4d53-9ecb-2803e2b90b70)| FunctionResult |修正ベッセル関数 Kn(x) を返します。|
|[BESSELY 関数](https://support.office.com/en-us/article/BESSELY-function-f3a356b3-da89-42c3-8974-2da54d6353a2)| FunctionResult |ベッセル関数 Yn(x) を返します。|
|[BETA.DIST 関数](https://support.office.com/en-us/article/BETADIST-function-11188c9c-780a-42c7-ba43-9ecb5a878d31)| FunctionResult |β 分布の累積分布関数の値を返します。|
|[BETA.INV 関数](https://support.office.com/en-us/article/BETAINV-function-e84cb8aa-8df0-4cf6-9892-83a341d252eb)| FunctionResult |指定された β 分布の累積分布関数の逆関数値を返します。|
|[BIN2DEC 関数](https://support.office.com/en-us/article/BIN2DEC-function-63905b57-b3a0-453d-99f4-647bb519cd6c)| FunctionResult |2 進数を 10 進数に変換します。|
|[BIN2HEX 関数](https://support.office.com/en-us/article/BIN2HEX-function-0375e507-f5e5-4077-9af8-28d84f9f41cc)| FunctionResult |2 進数を 16 進数に変換します。|
|[BIN2OCT 関数](https://support.office.com/en-us/article/BIN2OCT-function-0a4e01ba-ac8d-4158-9b29-16c25c4c23fd)| FunctionResult |2 進数を 8 進数に変換します。|
|[BINOM.DIST 関数](https://support.office.com/en-us/article/BINOMDIST-function-c5ae37b6-f39c-4be2-94c2-509a1480770c)| FunctionResult |二項分布の確率関数の値を返します。|
|[BINOM.DIST.RANGE 関数](https://support.office.com/en-us/article/BINOMDISTRANGE-function-17331329-74c7-4053-bb4c-6653a7421595)| FunctionResult |二項分布を使用した試行結果の確率を返します。|
|[BINOM.INV 関数](https://support.office.com/en-us/article/BINOMINV-function-80a0370c-ada6-49b4-83e7-05a91ba77ac9)| FunctionResult |累積二項分布の値が基準値以下になるような最小の値を返します。|
|[BITAND 関数](https://support.office.com/en-us/article/BITAND-function-8a2be3d7-91c3-4b48-9517-64548008563a)| FunctionResult |2 つの数値のビット演算 AND を返します。|
|[BITLSHIFT 関数](https://support.office.com/en-us/article/BITLSHIFT-function-c55bb27e-cacd-4c7c-b258-d80861a03c9c)| FunctionResult |左に移動数ビット (shift_amount) 移動する数値を返します。|
|[BITOR 関数](https://support.office.com/en-us/article/BITOR-function-f6ead5c8-5b98-4c9e-9053-8ad5234919b2)| FunctionResult |2 つの数値のビット演算 OR を返します。|
|[BITRSHIFT 関数](https://support.office.com/en-us/article/BITRSHIFT-function-274d6996-f42c-4743-abdb-4ff95351222c)| FunctionResult |右に移動数ビット (shift_amount) 移動する数値を返します。|
|[BITXOR 関数](https://support.office.com/en-us/article/BITXOR-function-c81306a1-03f9-4e89-85ac-b86c3cba10e4)| FunctionResult |2 つの数値のビット演算 "排他的 OR" を返します。|
|[CEILING.MATH 関数](https://support.office.com/en-us/article/CEILINGMATH-function-80f95d2f-b499-4eee-9f16-f795a8e306c8)| FunctionResult |数値を最も近い整数、または基準値に最も近い倍数に切り上げます。|
|[CEILING.PRECISE 関数](https://support.office.com/en-us/article/CEILINGPRECISE-function-f366a774-527a-4c92-ba49-af0a196e66cb)| FunctionResult |数値を最も近い整数、または基準値に最も近い倍数に切り上げます。数値の符号に関係なく、切り上げます。|
|[CHAR 関数](https://support.office.com/en-us/article/CHAR-function-bbd249c8-b36e-4a91-8017-1c133f9b837a)| FunctionResult |コード番号で指定された文字を返します。|
|[CHISQ.DIST 関数](https://support.office.com/en-us/article/CHISQDIST-function-8486b05e-5c05-4942-a9ea-f6b341518732)| FunctionResult |累積 β 確率密度関数の値を返します。|
|[CHISQ.DIST.RT 関数](https://support.office.com/en-us/article/CHISQDISTRT-function-dc4832e8-ed2b-49ae-8d7c-b28d5804c0f2)| FunctionResult |カイ 2 乗分布の片側確率の値を返します。|
|[CHISQ.INV 関数](https://support.office.com/en-us/article/CHISQINV-function-400db556-62b3-472d-80b3-254723e7092f)| FunctionResult |累積 β 確率密度関数の値を返します。|
|[CHISQ.INV.RT 関数](https://support.office.com/en-us/article/CHISQINVRT-function-435b5ed8-98d5-4da6-823f-293e2cbc94fe)| FunctionResult |カイ 2 乗分布の片側確率の逆関数値を返します。|
|[CHOOSE 関数](https://support.office.com/en-us/article/CHOOSE-function-fc5c184f-cb62-4ec7-a46e-38653b98f5bc)| FunctionResult |値のリストから値を選択します。|
|[CLEAN 関数](https://support.office.com/en-us/article/CLEAN-function-26f3d7c5-475f-4a9c-90e5-4b8ba987ba41)| FunctionResult |印刷できない文字を文字列から削除します。|
|[CODE 関数](https://support.office.com/en-us/article/CODE-function-c32b692b-2ed0-4a04-bdd9-75640144b928)| FunctionResult |テキスト文字列内の先頭文字の数値コードを返します。|
|[COLUMNS 関数](https://support.office.com/en-us/article/COLUMNS-function-4e8e7b4e-e603-43e8-b177-956088fa48ca)| FunctionResult |指定の範囲に含まれる列数を返します。|
|[COMBIN 関数](https://support.office.com/en-us/article/COMBIN-function-12a3f276-0a21-423a-8de6-06990aaf638a)| FunctionResult |指定された個数のオブジェクトを選択するときの組み合わせの数を返します。|
|[COMBINA 関数](https://support.office.com/en-us/article/COMBINA-function-efb49eaa-4f4c-4cd2-8179-0ddfcf9d035d)| FunctionResult |指定された個数の項目を選択するときの組み合わせ (反復あり) の数を返します|
|[COMPLEX 関数](https://support.office.com/en-us/article/COMPLEX-function-f0b8f3a9-51cc-4d6d-86fb-3a9362fa4128)| FunctionResult |実数係数と虚数係数を、複素数に変換します。|
|[CONCATENATE 関数](https://support.office.com/en-us/article/CONCATENATE-function-8f8ae884-2ca8-4f7a-b093-75d702bea31d)| FunctionResult |複数の文字列を 1 つの文字列に結合します。|
|[CONFIDENCE.NORM 関数](https://support.office.com/en-us/article/CONFIDENCENORM-function-7cec58a6-85bb-488d-91c3-63828d4fbfd4)| FunctionResult |母集団の平均に対する信頼区間を返します。|
|[CONFIDENCE.T 関数](https://support.office.com/en-us/article/CONFIDENCET-function-e8eca395-6c3a-4ba9-9003-79ccc61d3c53)| FunctionResult |スチューデントの t 分布を使用して、母集団の平均に対する信頼区間を返します。|
|[CONVERT 関数](https://support.office.com/en-us/article/CONVERT-function-d785bef1-808e-4aac-bdcd-666c810f9af2)| FunctionResult |数値の単位を変換します。|
|[COS 関数](https://support.office.com/en-us/article/COS-function-0fb808a5-95d6-4553-8148-22aebdce5f05)| FunctionResult |数値の余弦 (コサイン) を返します。|
|[COSH 関数](https://support.office.com/en-us/article/COSH-function-e460d426-c471-43e8-9540-a57ff3b70555)| FunctionResult |数値の双曲線余弦を返します。|
|[COT 関数](https://support.office.com/en-us/article/COT-function-c446f34d-6fe4-40dc-84f8-cf59e5f5e31a)| FunctionResult |角度のコタンジェントを返します。|
|[COTH 関数](https://support.office.com/en-us/article/COTH-function-2e0b4cb6-0ba0-403e-aed4-deaa71b49df5)| FunctionResult |双曲線余接 (ハイパーボリック コタンジェント) を返します。|
|[COUNT 関数](https://support.office.com/en-us/article/COUNT-function-a59cd7fc-b623-4d93-87a4-d23bf411294c)| FunctionResult |引数リストに含まれる数値の個数をカウントします。|
|[COUNTA 関数](https://support.office.com/en-us/article/COUNTA-function-7dc98875-d5c1-46f1-9a82-53f3219e2509)| FunctionResult |引数リストに含まれる値の個数をカウントします。|
|[COUNTBLANK 関数](https://support.office.com/en-us/article/COUNTBLANK-function-6a92d772-675c-4bee-b346-24af6bd3ac22)| FunctionResult |指定された範囲に含まれる空白セルの個数をカウントします。|
|[COUNTIF 関数](https://support.office.com/en-us/article/COUNTIF-function-e0de10c6-f885-4e71-abb4-1f464816df34)| FunctionResult |指定された範囲に含まれるセルのうち、検索条件に一致するセルの個数をカウントします。|
|[COUNTIFS 関数](https://support.office.com/en-us/article/COUNTIFS-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842)| FunctionResult |指定された範囲に含まれるセルのうち、複数の検索条件に一致するセルの個数を返します。|
|[COUPDAYBS 関数](https://support.office.com/en-us/article/COUPDAYBS-function-eb9a8dfb-2fb2-4c61-8e5d-690b320cf872)| FunctionResult |利払期間の第 1 日目から受渡日までの日数を返します。|
|[COUPDAYS 関数](https://support.office.com/en-us/article/COUPDAYS-function-cc64380b-315b-4e7b-950c-b30b0a76f671)| FunctionResult |受渡日を含む利払期間内の日数を返します。|
|[COUPDAYSNC 関数](https://support.office.com/en-us/article/COUPDAYSNC-function-5ab3f0b2-029f-4a8b-bb65-47d525eea547)| FunctionResult |受渡日から次の利払日までの日数を返します。|
|[COUPNCD 関数](https://support.office.com/en-us/article/COUPNCD-function-fd962fef-506b-4d9d-8590-16df5393691f)| FunctionResult |受渡日後の次の利払日を返します。|
|[COUPNUM 関数](https://support.office.com/en-us/article/COUPNUM-function-a90af57b-de53-4969-9c99-dd6139db2522)| FunctionResult |受渡日と満期日の間の利払回数を返します。|
|[COUPPCD 関数](https://support.office.com/en-us/article/COUPPCD-function-2eb50473-6ee9-4052-a206-77a9a385d5b3)| FunctionResult |受渡日の直前の利払日を返します。|
|[CSC 関数](https://support.office.com/en-us/article/CSC-function-07379361-219a-4398-8675-07ddc4f135c1)| FunctionResult |角度の余割 (コセカント) を返します。|
|[CSCH 関数](https://support.office.com/en-us/article/CSCH-function-f58f2c22-eb75-4dd6-84f4-a503527f8eeb)| FunctionResult |角度の双曲線余割を返します。|
|[CUMIPMT 関数](https://support.office.com/en-us/article/CUMIPMT-function-61067bb0-9016-427d-b95b-1a752af0e606)| FunctionResult |指定の期間に支払われる利息の累計を返します。|
|[CUMPRINC 関数](https://support.office.com/en-us/article/CUMPRINC-function-94a4516d-bd65-41a1-bc16-053a6af4c04d)| FunctionResult |指定期間に、貸付金に対して支払われる元金の累計を返します。|
|[DATE 関数](https://support.office.com/en-us/article/DATE-function-e36c0c8c-4104-49da-ab83-82328b832349)| FunctionResult |指定された日付に対応するシリアル値を返します。|
|[DATEVALUE 関数](https://support.office.com/en-us/article/DATEVALUE-function-df8b07d4-7761-4a93-bc33-b7471bbff252)| FunctionResult |日付を表す文字列をシリアル値に変換します。|
|[DAVERAGE 関数](https://support.office.com/en-us/article/DAVERAGE-function-a6a2d5ac-4b4b-48cd-a1d8-7b37834e5aee)| FunctionResult |選択したデータベース レコードの平均値を返します。|
|[DAY 関数](https://support.office.com/en-us/article/DAY-function-8a7d1cbb-6c7d-4ba1-8aea-25c134d03101)| FunctionResult |シリアル値を日付に変換します。|
|[DAYS 関数](https://support.office.com/en-us/article/DAYS-function-57740535-d549-4395-8728-0f07bff0b9df)| FunctionResult |2 つの日付間の日数を返します。|
|[DAYS360 関数](https://support.office.com/en-us/article/DAYS360-function-b9a509fd-49ef-407e-94df-0cbda5718c2a)| FunctionResult |1 年を 360 日として、2 つの日付間の日数を返します。|
|[DB 関数](https://support.office.com/en-us/article/DB-function-354e7d28-5f93-4ff1-8a52-eb4ee549d9d7)| FunctionResult |定率法 (Fixed-declining Balance Method) を利用して、特定の期における資産の減価償却費を返します。|
|[DBCS 関数](https://support.office.com/en-us/article/DBCS-function-a4025e73-63d2-4958-9423-21a24794c9e5)| FunctionResult |文字列内の半角 (1 バイト) の英数カナ文字を全角 (2 バイト) の文字に変換します。|
|[DCOUNT 関数](https://support.office.com/en-us/article/DCOUNT-function-c1fc7b93-fb0d-4d8d-97db-8d5f076eaeb1)| FunctionResult |データベース内にある数値を含むセルの個数をカウントします。|
|[DCOUNTA 関数](https://support.office.com/en-us/article/DCOUNTA-function-00232a6d-5a66-4a01-a25b-c1653fda1244)| FunctionResult |データベース内にある空白でないセルの個数をカウントします。|
|[DDB 関数](https://support.office.com/en-us/article/DDB-function-519a7a37-8772-4c96-85c0-ed2c209717a5)| FunctionResult |倍額定率法 (Double-declining Balance Method) または指定した他の方法を使用して、特定の期における資産の減価償却費を返します。|
|[DEC2BIN 関数](https://support.office.com/en-us/article/DEC2BIN-function-0f63dd0e-5d1a-42d8-b511-5bf5c6d43838)| FunctionResult |10 進数を 2 進数に変換します。|
|[DEC2HEX 関数](https://support.office.com/en-us/article/DEC2HEX-function-6344ee8b-b6b5-4c6a-a672-f64666704619)| FunctionResult |10 進数を 16 進数に変換します。|
|[DEC2OCT 関数](https://support.office.com/en-us/article/DEC2OCT-function-c9d835ca-20b7-40c4-8a9e-d3be351ce00f)| FunctionResult |10 進数を 8 進数に変換します。|
|[DECIMAL 関数](https://support.office.com/en-us/article/DECIMAL-function-ee554665-6176-46ef-82de-0a283658da2e)| FunctionResult |指定された底の数値のテキスト表現を 10 進数に変換します。|
|[DEGREES 関数](https://support.office.com/en-us/article/DEGREES-function-4d6ec4db-e694-4b94-ace0-1cc3f61f9ba1)| FunctionResult |ラジアンを度に変換します。|
|[DELTA 関数](https://support.office.com/en-us/article/DELTA-function-2f763672-c959-4e07-ac33-fe03220ba432)| FunctionResult |2 つの値が等しいかどうかをテストします。|
|[DEVSQ 関数](https://support.office.com/en-us/article/DEVSQ-function-8b739616-8376-4df5-8bd0-cfe0a6caf444)| FunctionResult |偏差の平方和を返します。|
|[DGET 関数](https://support.office.com/en-us/article/DGET-function-455568bf-4eef-45f7-90f0-ec250d00892e)| FunctionResult |指定された条件に一致する 1 つのレコードをデータベースから抽出します。|
|[DISC 関数](https://support.office.com/en-us/article/DISC-function-71fce9f3-3f05-4acf-a5a3-eac6ef4daa53)| FunctionResult |証券に対する割引率を返します。|
|[DMAX 関数](https://support.office.com/en-us/article/DMAX-function-f4e8209d-8958-4c3d-a1ee-6351665d41c2)| FunctionResult |選択したデータベース レコードの最大値を返します。|
|[DMIN 関数](https://support.office.com/en-us/article/DMIN-function-4ae6f1d9-1f26-40f1-a783-6dc3680192a3)| FunctionResult |選択したデータベース レコードの最小値を返します。|
|[DOLLAR 関数](https://support.office.com/en-us/article/DOLLAR-function-a6cd05d9-9740-4ad3-a469-8109d18ff611)| FunctionResult |ドル ($) 通貨書式を使用して、数値を文字列に変換します。|
|[DOLLARDE 関数](https://support.office.com/en-us/article/DOLLARDE-function-db85aab0-1677-428a-9dfd-a38476693427)| FunctionResult |分数で表されたドル単位の価格を、小数表示のドル価格に変換します。|
|[DOLLARFR 関数](https://support.office.com/en-us/article/DOLLARFR-function-0835d163-3023-4a33-9824-3042c5d4f495)| FunctionResult |小数で表されたドル単位の価格を、分数表示のドル価格に変換します。|
|[DPRODUCT 関数](https://support.office.com/en-us/article/DPRODUCT-function-4f96b13e-d49c-47a7-b769-22f6d017cb31)| FunctionResult |データベース内の、条件に一致するレコードの特定のフィールド値を乗算します。|
|[DSTDEV 関数](https://support.office.com/en-us/article/DSTDEV-function-026b8c73-616d-4b5e-b072-241871c4ab96)| FunctionResult |選択したデータベース レコードの標本に基づいて、標準偏差の推定値を返します。|
|[DSTDEVP 関数](https://support.office.com/en-us/article/DSTDEVP-function-04b78995-da03-4813-bbd9-d74fd0f5d94b)| FunctionResult |選択したデータベース レコードの母集団全体に基づいて標準偏差を算出します。|
|[DSUM 関数](https://support.office.com/en-us/article/DSUM-function-53181285-0c4b-4f5a-aaa3-529a322be41b)| FunctionResult |データベース内の、条件に一致するレコードのフィールド列にある数値を合計します。|
|[DURATION 関数](https://support.office.com/en-us/article/DURATION-function-b254ea57-eadc-4602-a86a-c8e369334038)| FunctionResult |定期的に利子が支払われる証券の年間のマコーレー デュレーションを返します。|
|[DVAR 関数](https://support.office.com/en-us/article/DVAR-function-d6747ca9-99c7-48bb-996e-9d7af00f3ed1)| FunctionResult |選択したデータベース レコードの標本に基づく分散の推定値を返します。|
|[DVARP 関数](https://support.office.com/en-us/article/DVARP-function-eb0ba387-9cb7-45c8-81e9-0394912502fc)| FunctionResult |選択したデータベース レコードの母集団全体に基づく分散を算出します。|
|[EDATE 関数](https://support.office.com/en-us/article/EDATE-function-3c920eb2-6e66-44e7-a1f5-753ae47ee4f5)| FunctionResult |開始日から起算して、指定した月数だけ前または後の日付に対応するシリアル値を返します。|
|[EFFECT 関数](https://support.office.com/en-us/article/EFFECT-function-910d4e4c-79e2-4009-95e6-507e04f11bc4)| FunctionResult |実効年利率を返します。|
|[EOMONTH 関数](https://support.office.com/en-us/article/EOMONTH-function-7314ffa1-2bc9-4005-9d66-f49db127d628)| FunctionResult |指定した月数だけ前または後の月の最終日に対応するシリアル値を返します。|
|[ERF 関数](https://support.office.com/en-us/article/ERF-function-c53c7e7b-5482-4b6c-883e-56df3c9af349)| FunctionResult |誤差関数の値を返します。|
|[ERF.PRECISE 関数](https://support.office.com/en-us/article/ERFPRECISE-function-9a349593-705c-4278-9a98-e4122831a8e0)| FunctionResult |誤差関数の値を返します。|
|[ERFC 関数](https://support.office.com/en-us/article/ERFC-function-736e0318-70ba-4e8b-8d08-461fe68b71b3)| FunctionResult |相補誤差関数の値を返します。|
|[ERFC.PRECISE 関数](https://support.office.com/en-us/article/ERFCPRECISE-function-e90e6bab-f45e-45df-b2ac-cd2eb4d4a273)| FunctionResult |x から無限大の範囲で、相補誤差関数の積分値を返します。|
|[ERROR.TYPE 関数](https://support.office.com/en-us/article/ERRORTYPE-function-10958677-7c8d-44f7-ae77-b9a9ee6eefaa)| FunctionResult |エラーの種類に対応する数値を返します。|
|[EVEN 関数](https://support.office.com/en-us/article/EVEN-function-197b5f06-c795-4c1e-8696-3c3b8a646cf9)| FunctionResult |指定された数値を最も近い偶数に切り上げた値を返します。|
|[EXACT 関数](https://support.office.com/en-us/article/EXACT-function-d3087698-fc15-4a15-9631-12575cf29926)| FunctionResult |2 つのテキスト値が等しいかどうかを判定します。|
|[EXP 関数](https://support.office.com/en-us/article/EXP-function-c578f034-2c45-4c37-bc8c-329660a63abe)| FunctionResult |e を底とする数値のべき乗を返します。|
|[EXPON.DIST 関数](https://support.office.com/en-us/article/EXPONDIST-function-4c12ae24-e563-4155-bf3e-8b78b6ae140e)| FunctionResult |指数分布を返します。|
|[F.DIST 関数](https://support.office.com/en-us/article/FDIST-function-a887efdc-7c8e-46cb-a74a-f884cd29b25d)| FunctionResult |F 分布の確率関数の値を返します。|
|[F.DIST.RT 関数](https://support.office.com/en-us/article/FDISTRT-function-d74cbb00-6017-4ac9-b7d7-6049badc0520)| FunctionResult |F 分布の確率関数の値を返します。|
|[F.INV 関数](https://support.office.com/en-us/article/FINV-function-0dda0cf9-4ea0-42fd-8c3c-417a1ff30dbe)| FunctionResult |F 分布の確率関数の逆関数の値を返します。|
|[F.INV.RT 関数](https://support.office.com/en-us/article/FINVRT-function-d371aa8f-b0b1-40ef-9cc2-496f0693ac00)| FunctionResult |F 分布の確率関数の逆関数の値を返します。|
|[FACT 関数](https://support.office.com/en-us/article/FACT-function-ca8588c2-15f2-41c0-8e8c-c11bd471a4f3)| FunctionResult |数値の階乗を返します。|
|[FACTDOUBLE 関数](https://support.office.com/en-us/article/FACTDOUBLE-function-e67697ac-d214-48eb-b7b7-cce2589ecac8)| FunctionResult |数値の二重階乗を返します。|
|[FALSE 関数](https://support.office.com/en-us/article/FALSE-function-2d58dfa5-9c03-4259-bf8f-f0ae14346904)| FunctionResult |論理値 FALSE を返します。|
|[FIND 関数、FINDB 関数](https://support.office.com/en-us/article/FIND-FINDB-functions-c7912941-af2a-4bdf-a553-d0d89b0a0628)| FunctionResult |指定されたテキスト値を他のテキスト値の中で検索します。大文字と小文字は区別されます。|
|[FISHER 関数](https://support.office.com/en-us/article/FISHER-function-d656523c-5076-4f95-b87b-7741bf236c69)| FunctionResult |フィッシャー変換の値を返します。|
|[FISHERINV 関数](https://support.office.com/en-us/article/FISHERINV-function-62504b39-415a-4284-a285-19c8e82f86bb)| FunctionResult |フィッシャー変換の逆関数値を返します。|
|[FIXED 関数](https://support.office.com/en-us/article/FIXED-function-ffd5723c-324c-45e9-8b96-e41be2a8274a)| FunctionResult |数値を、一定の桁数のテキストとして書式設定します。|
|[FLOOR 関数](https://support.office.com/en-us/article/FLOOR-function-14bb497c-24f2-4e04-b327-b0b4de5a8886)| FunctionResult |数値を指定された桁数で切り捨てます。|
|[FLOOR.MATH 関数](https://support.office.com/en-us/article/FLOORMATH-function-c302b599-fbdb-4177-ba19-2c2b1249a2f5)| FunctionResult |最も近い整数値、または基準値の倍数のうちで最も近い値に切り下げます。|
|[FLOOR.PRECISE 関数](https://support.office.com/en-us/article/FLOORPRECISE-function-f769b468-1452-4617-8dc3-02f842a0702e)| FunctionResult |最も近い整数値、または基準値の倍数のうちで最も近い値に切り下げます。数値の符号に関係なく、切り下げます。|
|[FORECAST 関数](https://support.office.com/en-us/article/FORECAST-function-50ca49c9-7b40-4892-94e4-7ad38bbeda99)| FunctionResult |線形トレンドに沿った値を返します。|
|[FORECAST.ETS 関数](https://support.office.com/en-us/article/FORECASTETS-function-897a2fe9-6595-4680-a0b0-93e0308d5f6e)| FunctionResult |指数平滑化 (ETS) アルゴリズムの AAA バージョンを使って、既存の (履歴) 値に基づき将来の値を返します。|
|[FORECAST.ETS.CONFINT 関数](https://support.office.com/en-us/article/FORECASTETSCONFINT-function-897a2fe9-6595-4680-a0b0-93e0308d5f6e)| FunctionResult |指定した対象の日付における、予測値に対する信頼区間を返します。|
|[FORECAST.ETS.SEASONALITY 関数](https://support.office.com/en-us/article/FORECASTETSSEASONALITY-function-897a2fe9-6595-4680-a0b0-93e0308d5f6e)| FunctionResult |指定した時系列に関して Excel が検出した繰り返しパターンの長さを返します。|
|[FORECAST.ETS.STAT 関数](https://support.office.com/en-us/article/FORECASTETSSTAT-function-897a2fe9-6595-4680-a0b0-93e0308d5f6e)| FunctionResult |時系列予測の結果として統計値を返します。|
|[FORECAST.LINEAR 関数](https://support.office.com/en-us/article/FORECASTLINEAR-function-897a2fe9-6595-4680-a0b0-93e0308d5f6e)| FunctionResult |既存の値に基づいて将来値を返します。|
|[FV 関数](https://support.office.com/en-us/article/FV-function-2eef9f44-a084-4c61-bdd8-4fe4bb1b71b3)| FunctionResult |投資の将来価値を返します。|
|[FVSCHEDULE 関数](https://support.office.com/en-us/article/FVSCHEDULE-function-bec29522-bd87-4082-bab9-a241f3fb251d)| FunctionResult |一連の金利を複利計算することにより、初期投資した元金の将来の価値を返します。|
|[GAMMA 関数](https://support.office.com/en-us/article/GAMMA-function-ce1702b1-cf55-471d-8307-f83be0fc5297)| FunctionResult |Gamma 関数値を返します。|
|[GAMMA.DIST 関数](https://support.office.com/en-us/article/GAMMADIST-function-9b6f1538-d11c-4d5f-8966-21f6a2201def)| FunctionResult |ガンマ分布の値を返します。|
|[GAMMA.INV 関数](https://support.office.com/en-us/article/GAMMAINV-function-74991443-c2b0-4be5-aaab-1aa4d71fbb18)| FunctionResult |ガンマの累積分布の逆関数値を返します。|
|[GAMMALN 関数](https://support.office.com/en-us/article/GAMMALN-function-b838c48b-c65f-484f-9e1d-141c55470eb9)| FunctionResult |ガンマ関数 Γ(x) の値の自然対数を返します。|
|[GAMMALN.PRECISE 関数](https://support.office.com/en-us/article/GAMMALNPRECISE-function-5cdfe601-4e1e-4189-9d74-241ef1caa599)| FunctionResult |ガンマ関数 Γ(x) の値の自然対数を返します。|
|[GAUSS 関数](https://support.office.com/en-us/article/GAUSS-function-069f1b4e-7dee-4d6a-a71f-4b69044a6b33)| FunctionResult |標準正規分布の累積分布関数より 0.5 小さい値を返します。|
|[GCD 関数](https://support.office.com/en-us/article/GCD-function-d5107a51-69e3-461f-8e4c-ddfc21b5073a)| FunctionResult |最大公約数を返します。|
|[GEOMEAN 関数](https://support.office.com/en-us/article/GEOMEAN-function-db1ac48d-25a5-40a0-ab83-0b38980e40d5)| FunctionResult |相乗平均を返します。|
|[GESTEP 関数](https://support.office.com/en-us/article/GESTEP-function-f37e7d2a-41da-4129-be95-640883fca9df)| FunctionResult |数値がしきい値以上であるかどうかをテストします。|
|[HARMEAN 関数](https://support.office.com/en-us/article/HARMEAN-function-5efd9184-fab5-42f9-b1d3-57883a1d3bc6)| FunctionResult |調和平均を返します。|
|[HEX2BIN 関数](https://support.office.com/en-us/article/HEX2BIN-function-a13aafaa-5737-4920-8424-643e581828c1)| FunctionResult |16 進数を 2 進数に変換します。|
|[HEX2DEC 関数](https://support.office.com/en-us/article/HEX2DEC-function-8c8c3155-9f37-45a5-a3ee-ee5379ef106e)| FunctionResult |16 進数を 10 進数に変換します。|
|[HEX2OCT 関数](https://support.office.com/en-us/article/HEX2OCT-function-54d52808-5d19-4bd0-8a63-1096a5d11912)| FunctionResult |16 進数を 8 進数に変換します。|
|[HLOOKUP 関数](https://support.office.com/en-us/article/HLOOKUP-function-a3034eec-b719-4ba3-bb65-e1ad662ed95f)| FunctionResult |配列の上端行で特定の値を検索し、対応するセルの値を返します。|
|[HOUR 関数](https://support.office.com/en-us/article/HOUR-function-a3afa879-86cb-4339-b1b5-2dd2d7310ac7)| FunctionResult |シリアル値を時刻に変換します。|
|[HYPERLINK 関数](https://support.office.com/en-us/article/HYPERLINK-function-333c7ce6-c5ae-4164-9c47-7de9b76f577f)| FunctionResult |ネットワーク サーバー、イントラネット、またはインターネット上に格納されているドキュメントを開くショートカットまたはジャンプを作成します。|
|[HYPGEOM.DIST 関数](https://support.office.com/en-us/article/HYPGEOMDIST-function-6dbd547f-1d12-4b1f-8ae5-b0d9e3d22fbf)| FunctionResult |超幾何分布を返します。|
|[IF 関数](https://support.office.com/en-us/article/IF-function-69aed7c9-4e8a-4755-a9bc-aa8bbff73be2)| FunctionResult |実行する論理テストを指定します。|
|[IMABS 関数](https://support.office.com/en-us/article/IMABS-function-b31e73c6-d90c-4062-90bc-8eb351d765a1)| FunctionResult |指定した複素数の絶対値を返します。|
|[IMAGINARY 関数](https://support.office.com/en-us/article/IMAGINARY-function-dd5952fd-473d-44d9-95a1-9a17b23e428a)| FunctionResult |指定した複素数の虚数係数を返します。|
|[IMARGUMENT 関数](https://support.office.com/en-us/article/IMARGUMENT-function-eed37ec1-23b3-4f59-b9f3-d340358a034a)| FunctionResult |偏角シータを (ラジアンで表した角度で) 返します。|
|[IMCONJUGATE 関数](https://support.office.com/en-us/article/IMCONJUGATE-function-2e2fc1ea-f32b-4f9b-9de6-233853bafd42)| FunctionResult |複素数の複素共役を返します。|
|[IMCOS 関数](https://support.office.com/en-us/article/IMCOS-function-dad75277-f592-4a6b-ad6c-be93a808a53c)| FunctionResult |複素数のコサインを返します。|
|[IMCOSH 関数](https://support.office.com/en-us/article/IMCOSH-function-053e4ddb-4122-458b-be9a-457c405e90ff)| FunctionResult |複素数の双曲線余弦を返します。|
|[IMCOT 関数](https://support.office.com/en-us/article/IMCOT-function-dc6a3607-d26a-4d06-8b41-8931da36442c)| FunctionResult |複素数の余接 (コタンジェント) を返します。|
|[IMCSC 関数](https://support.office.com/en-us/article/IMCSC-function-9e158d8f-2ddf-46cd-9b1d-98e29904a323)| FunctionResult |複素数の余割 (コセカント) を返します。|
|[IMCSCH 関数](https://support.office.com/en-us/article/IMCSCH-function-c0ae4f54-5f09-4fef-8da0-dc33ea2c5ca9)| FunctionResult |複素数の双曲線余割を返します。|
|[IMDIV 関数](https://support.office.com/en-us/article/IMDIV-function-a505aff7-af8a-4451-8142-77ec3d74d83f)| FunctionResult |2 つの複素数の商を返します。|
|[IMEXP 関数](https://support.office.com/en-us/article/IMEXP-function-c6f8da1f-e024-4c0c-b802-a60e7147a95f)| FunctionResult |複素数のべき乗を返します。|
|[IMLN 関数](https://support.office.com/en-us/article/IMLN-function-32b98bcf-8b81-437c-a636-6fb3aad509d8)| FunctionResult |複素数の自然対数を返します。|
|[IMLOG10 関数](https://support.office.com/en-us/article/IMLOG10-function-58200fca-e2a2-4271-8a98-ccd4360213a5)| FunctionResult |複素数の 10 を底とする対数 (常用対数) を返します。|
|[IMLOG2 関数](https://support.office.com/en-us/article/IMLOG2-function-152e13b4-bc79-486c-a243-e6a676878c51)| FunctionResult |複素数の 2 を底とする対数を返します。|
|[IMPOWER 関数](https://support.office.com/en-us/article/IMPOWER-function-210fd2f5-f8ff-4c6a-9d60-30e34fbdef39)| FunctionResult |複素数の整数乗を返します。|
|[IMPRODUCT 関数](https://support.office.com/en-us/article/IMPRODUCT-function-2fb8651a-a4f2-444f-975e-8ba7aab3a5ba)| FunctionResult |2 から 255 個の複素数の積を返します。|
|[IMREAL 関数](https://support.office.com/en-us/article/IMREAL-function-d12bc4c0-25d0-4bb3-a25f-ece1938bf366)| FunctionResult |複素数の実数係数を返します。|
|[IMSEC 関数](https://support.office.com/en-us/article/IMSEC-function-6df11132-4411-4df4-a3dc-1f17372459e0)| FunctionResult |複素数の正割 (セカント) を返します。|
|[IMSECH 関数](https://support.office.com/en-us/article/IMSECH-function-f250304f-788b-4505-954e-eb01fa50903b)| FunctionResult |複素数の双曲線正割を返します。|
|[IMSIN 関数](https://support.office.com/en-us/article/IMSIN-function-1ab02a39-a721-48de-82ef-f52bf37859f6)| FunctionResult |複素数の正弦を返します。|
|[IMSINH 関数](https://support.office.com/en-us/article/IMSINH-function-dfb9ec9e-8783-4985-8c42-b028e9e8da3d)| FunctionResult |複素数の双曲線正弦を返します。|
|[IMSQRT 関数](https://support.office.com/en-us/article/IMSQRT-function-e1753f80-ba11-4664-a10e-e17368396b70)| FunctionResult |複素数の平方根を返します。|
|[IMSUB 関数](https://support.office.com/en-us/article/IMSUB-function-2e404b4d-4935-4e85-9f52-cb08b9a45054)| FunctionResult |2 つの複素数の差を返します。|
|[IMSUM 関数](https://support.office.com/en-us/article/IMSUM-function-81542999-5f1c-4da6-9ffe-f1d7aaa9457f)| FunctionResult |複素数の和を返します。|
|[IMTAN 関数](https://support.office.com/en-us/article/IMTAN-function-8478f45d-610a-43cf-8544-9fc0b553a132)| FunctionResult |複素数の正接 (タンジェント) を返します。|
|[INT 関数](https://support.office.com/en-us/article/INT-function-a6c4af9e-356d-4369-ab6a-cb1fd9d343ef)| FunctionResult |指定された数値を最も近い整数に切り捨てます。|
|[INTRATE 関数](https://support.office.com/en-us/article/INTRATE-function-5cb34dde-a221-4cb6-b3eb-0b9e55e1316f)| FunctionResult |全額投資された証券の利率を返します。|
|[IPMT 関数](https://support.office.com/en-us/article/IPMT-function-5cce0ad6-8402-4a41-8d29-61a0b054cb6f)| FunctionResult |投資の指定された期に支払われる金利を返します。|
|[IRR 関数](https://support.office.com/en-us/article/IRR-function-64925eaa-9988-495b-b290-3ad0c163c1bc)| FunctionResult |一連のキャッシュ フローに対する内部利益率を返します。|
|[ISERR 関数](https://support.office.com/en-us/article/ISERR-function-0f2d7971-6019-40a0-a171-f2d869135665)| FunctionResult |対象が #N/A 以外のエラー値のときに TRUE を返します。|
|[ISERROR 関数](https://support.office.com/en-us/article/ISERROR-function-0f2d7971-6019-40a0-a171-f2d869135665)| FunctionResult |対象が任意のエラー値のときに TRUE を返します。|
|[ISEVEN 関数](https://support.office.com/en-us/article/ISEVEN-function-aa15929a-d77b-4fbb-92f4-2f479af55356)| FunctionResult |数値が偶数のときに TRUE を返します|
|[ISFORMULA 関数](https://support.office.com/en-us/article/ISFORMULA-function-e4d1355f-7121-4ef2-801e-3839bfd6b1e5)| FunctionResult |数式が含まれるセルへの参照がある場合に TRUE を返します。|
|[ISLOGICAL 関数](https://support.office.com/en-us/article/ISLOGICAL-function-0f2d7971-6019-40a0-a171-f2d869135665)| FunctionResult |対象が論理値のときに TRUE を返します。|
|[ISNA 関数](https://support.office.com/en-us/article/ISNA-function-0f2d7971-6019-40a0-a171-f2d869135665)| FunctionResult |対象がエラー値 #N/A のときに TRUE を返します。|
|[ISNONTEXT 関数](https://support.office.com/en-us/article/ISNONTEXT-function-0f2d7971-6019-40a0-a171-f2d869135665)| FunctionResult |対象が文字列以外のときに TRUE を返します。|
|[ISNUMBER 関数](https://support.office.com/en-us/article/ISNUMBER-function-0f2d7971-6019-40a0-a171-f2d869135665)| FunctionResult |対象が数値のときに TRUE を返します。|
|[ISO.CEILING 関数](https://support.office.com/en-us/article/ISOCEILING-function-e587bb73-6cc2-4113-b664-ff5b09859a83)| FunctionResult |最も近い整数に切り上げた値、または、指定された基準値の倍数のうち最も近い値を返します。|
|[ISODD 関数](https://support.office.com/en-us/article/ISODD-function-0f2d7971-6019-40a0-a171-f2d869135665)| FunctionResult |数値が奇数のときに TRUE を返します|
|[ISOWEEKNUM 関数](https://support.office.com/en-us/article/ISOWEEKNUM-function-1c2d0afe-d25b-4ab1-8894-8d0520e90e0e)| FunctionResult |指定された日付のその年における ISO 週番号を返します。|
|[ISPMT 関数](https://support.office.com/en-us/article/ISPMT-function-fa58adb6-9d39-4ce0-8f43-75399cea56cc)| FunctionResult |投資の指定された期に支払われる金利を計算します。|
|[ISREF 関数](https://support.office.com/en-us/article/ISREF-function-0f2d7971-6019-40a0-a171-f2d869135665)| FunctionResult |対象が参照であるときに TRUE を返します。|
|[ISTEXT 関数](https://support.office.com/en-us/article/ISTEXT-function-0f2d7971-6019-40a0-a171-f2d869135665)| FunctionResult |対象がテキストであるときに TRUE を返します。|
|[KURT 関数](https://support.office.com/en-us/article/KURT-function-bc3a265c-5da4-4dcb-b7fd-c237789095ab)| FunctionResult |データ セットの尖度を返します。|
|[LARGE 関数](https://support.office.com/en-us/article/LARGE-function-3af0af19-1190-42bb-bb8b-01672ec00a64)| FunctionResult |指定されたデータ セットの中で k 番目に大きなデータを返します。|
|[LCM 関数](https://support.office.com/en-us/article/LCM-function-7152b67a-8bb5-4075-ae5c-06ede5563c94)| FunctionResult |最小公倍数を返します。|
|[LEFT 関数、LEFTB 関数](https://support.office.com/en-us/article/LEFT-LEFTB-functions-9203d2d2-7960-479b-84c6-1ea52b99640c)| FunctionResult |文字列の先頭 (左端) から指定された文字数の文字を返します。|
|[LEN 関数、LENB 関数](https://support.office.com/en-us/article/LEN-LENB-functions-29236f94-cedc-429d-affd-b5e33d2c67cb)| FunctionResult |文字列に含まれる文字数を返します。|
|[LN 関数](https://support.office.com/en-us/article/LN-function-81fe1ed7-dac9-4acd-ba1d-07a142c6118f)| FunctionResult |数値の自然対数を返します。|
|[LOG 関数](https://support.office.com/en-us/article/LOG-function-4e82f196-1ca9-4747-8fb0-6c4a3abb3280)| FunctionResult |指定された数を底とする数値の対数を返します。|
|[LOG10 関数](https://support.office.com/en-us/article/LOG10-function-c75b881b-49dd-44fb-b6f4-37e3486a0211)| FunctionResult |10 を底とする数値の対数 (常用対数) を返します。|
|[LOGNORM.DIST 関数](https://support.office.com/en-us/article/LOGNORMDIST-function-eb60d00b-48a9-4217-be2b-6074aee6b070)| FunctionResult |対数の累積分布の値を返します。|
|[LOGNORM.INV 関数](https://support.office.com/en-us/article/LOGNORMINV-function-fe79751a-f1f2-4af8-a0a1-e151b2d4f600)| FunctionResult |対数の累積分布の逆関数値を返します。|
|[LOOKUP 関数](https://support.office.com/en-us/article/LOOKUP-function-446d94af-663b-451d-8251-369d5e3864cb)| FunctionResult |ベクトルまたは配列を検索して、対応する値を返します。|
|[LOWER 関数](https://support.office.com/en-us/article/LOWER-function-3f21df02-a80c-44b2-afaf-81358f9fdeb4)| FunctionResult |テキストを小文字に変換します。|
|[MATCH 関数](https://support.office.com/en-us/article/MATCH-function-e8dffd45-c762-47d6-bf89-533f4a37673a)| FunctionResult |参照または配列で値を検索します。|
|[MAX 関数](https://support.office.com/en-us/article/MAX-function-e0012414-9ac8-4b34-9a47-73e662c08098)| FunctionResult |引数リストに含まれる最大値を返します。|
|[MAXA 関数](https://support.office.com/en-us/article/MAXA-function-814bda1e-3840-4bff-9365-2f59ac2ee62d)| FunctionResult |数値、文字列、および論理値を含む引数リストから最大値を返します。|
|[MDURATION 関数](https://support.office.com/en-us/article/MDURATION-function-b3786a69-4f20-469a-94ad-33e5b90a763c)| FunctionResult |額面価格を $100 と仮定して、証券に対する修正済マコーレー デュレーションを返します。|
|[MEDIAN 関数](https://support.office.com/en-us/article/MEDIAN-function-d0916313-4753-414c-8537-ce85bdd967d2)| FunctionResult |指定された数値のメジアン (中央値) を返します。|
|[MID 関数、MIDB 関数](https://support.office.com/en-us/article/MID-MIDB-functions-d5f9e25c-d7d6-472e-b568-4ecb12433028)| FunctionResult |文字列の任意の位置から指定された文字数の文字を返します。|
|[MIN 関数](https://support.office.com/en-us/article/MIN-function-61635d12-920f-4ce2-a70f-96f202dcc152)| FunctionResult |引数リストに含まれる最小値を返します。|
|[MINA 関数](https://support.office.com/en-us/article/MINA-function-245a6f46-7ca5-4dc7-ab49-805341bc31d3)| FunctionResult |数値、文字列、および論理値を含む引数リストから最小値を返します。|
|[MINUTE 関数](https://support.office.com/en-us/article/MINUTE-function-af728df0-05c4-4b07-9eed-a84801a60589)| FunctionResult |シリアル値を時刻の分に変換します。|
|[MIRR 関数](https://support.office.com/en-us/article/MIRR-function-b020f038-7492-4fb4-93c1-35c345b53524)| FunctionResult |支払い (負の値) と収益 (正の値) のキャッシュ フローがさまざまな率で行われる場合の修正内部利益率を返します。|
|[MOD 関数](https://support.office.com/en-us/article/MOD-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3)| FunctionResult |除算の剰余を返します。|
|[MONTH 関数](https://support.office.com/en-us/article/MONTH-function-579a2881-199b-48b2-ab90-ddba0eba86e8)| FunctionResult |シリアル値を月に変換します。|
|[MROUND 関数](https://support.office.com/en-us/article/MROUND-function-c299c3b0-15a5-426d-aa4b-d2d5b3baf427)| FunctionResult |指定された値の倍数になるように、数値を四捨五入します。|
|[MULTINOMIAL 関数](https://support.office.com/en-us/article/MULTINOMIAL-function-6fa6373c-6533-41a2-a45e-a56db1db1bf6)| FunctionResult |指定された複数の数値の多項係数を返します。|
|[N 関数](https://support.office.com/en-us/article/N-function-a624cad1-3635-4208-b54a-29733d1278c9)| FunctionResult |値を数値に変換します。|
|[NA 関数](https://support.office.com/en-us/article/NA-function-5469c2d1-a90c-4fb5-9bbc-64bd9bb6b47c)| FunctionResult |エラー値 #N/A を返します。|
|[NEGBINOM.DIST 関数](https://support.office.com/en-us/article/NEGBINOMDIST-function-c8239f89-c2d0-45bd-b6af-172e570f8599)| FunctionResult |負の二項分布を返します。|
|[NETWORKDAYS 関数](https://support.office.com/en-us/article/NETWORKDAYS-function-48e717bf-a7a3-495f-969e-5005e3eb18e7)| FunctionResult |2 つの日付間の稼働日の日数を返します。|
|[NETWORKDAYS.INTL 関数](https://support.office.com/en-us/article/NETWORKDAYSINTL-function-a9b26239-4f20-46a1-9ab8-4e925bfd5e28)| FunctionResult |週末がどの曜日で何日間あるかを示すパラメーターを使用して、2 つの日付間にある稼働日の日数を返します。|
|[NOMINAL 関数](https://support.office.com/en-us/article/NOMINAL-function-7f1ae29b-6b92-435e-b950-ad8b190ddd2b)| FunctionResult |名目年利率を返します。|
|[NORM.DIST 関数](https://support.office.com/en-us/article/NORMDIST-function-edb1cc14-a21c-4e53-839d-8082074c9f8d)| FunctionResult |正規分布の累積分布の値を返します。|
|[NORM.INV 関数](https://support.office.com/en-us/article/NORMINV-function-54b30935-fee7-493c-bedb-2278a9db7e13)| FunctionResult |正規分布の累積分布の逆関数値を返します。|
|[NORM.S.DIST 関数](https://support.office.com/en-us/article/NORMSDIST-function-1e787282-3832-4520-a9ae-bd2a8d99ba88)| FunctionResult |標準正規分布の累積分布の値を返します。|
|[NORM.S.INV 関数](https://support.office.com/en-us/article/NORMSINV-function-d6d556b4-ab7f-49cd-b526-5a20918452b1)| FunctionResult |標準正規分布の累積分布の逆関数値を返します。|
|[NOT 関数](https://support.office.com/en-us/article/NOT-function-9cfc6011-a054-40c7-a140-cd4ba2d87d77)| FunctionResult |引数の論理値を逆にして返します。|
|[NOW 関数](https://support.office.com/en-us/article/NOW-function-3337fd29-145a-4347-b2e6-20c904739c46)| FunctionResult |現在の日付と時刻に対応するシリアル値を返します。|
|[NPER 関数](https://support.office.com/en-us/article/NPER-function-240535b5-6653-4d2d-bfcf-b6a38151d815)| FunctionResult |投資に必要な期間を返します。|
|[NPV 関数](https://support.office.com/en-us/article/NPV-function-8672cb67-2576-4d07-b67b-ac28acf2a568)| FunctionResult |定期的に発生する一連のキャッシュ フローと割引率に基づいて、投資の正味現在価値を返します。|
|[NUMBERVALUE 関数](https://support.office.com/en-us/article/NUMBERVALUE-function-1b05c8cf-2bfa-4437-af70-596c7ea7d879)| FunctionResult |文字列をロケールに依存しない方法で数値に変換します。|
|[OCT2BIN 関数](https://support.office.com/en-us/article/OCT2BIN-function-55383471-3c56-4d27-9522-1a8ec646c589)| FunctionResult |8 進数を 2 進数に変換します。|
|[OCT2DEC 関数](https://support.office.com/en-us/article/OCT2DEC-function-87606014-cb98-44b2-8dbb-e48f8ced1554)| FunctionResult |8 進数を 10 進数に変換します。|
|[OCT2HEX 関数](https://support.office.com/en-us/article/OCT2HEX-function-912175b4-d497-41b4-a029-221f051b858f)| FunctionResult |8 進数を 16 進数に変換します。|
|[ODD 関数](https://support.office.com/en-us/article/ODD-function-deae64eb-e08a-4c88-8b40-6d0b42575c98)| FunctionResult |指定された数値を最も近い奇数に切り上げた値を返します。|
|[ODDFPRICE 関数](https://support.office.com/en-us/article/ODDFPRICE-function-d7d664a8-34df-4233-8d2b-922bcf6a69e1)| FunctionResult |1 期目の日数が半端な証券に対して、額面 $100 あたりの価格を返します。|
|[ODDFYIELD 関数](https://support.office.com/en-us/article/ODDFYIELD-function-66bc8b7b-6501-4c93-9ce3-2fd16220fe37)| FunctionResult |1 期目の日数が半端な証券の利回りを返します。|
|[ODDLPRICE 関数](https://support.office.com/en-us/article/ODDLPRICE-function-fb657749-d200-4902-afaf-ed5445027fc4)| FunctionResult |最終期の日数が半端な証券に対して、額面 $100 あたりの価格を返します。|
|[ODDLYIELD 関数](https://support.office.com/en-us/article/ODDLYIELD-function-c873d088-cf40-435f-8d41-c8232fee9238)| FunctionResult |最終期の日数が半端な証券の利回りを返します。|
|[OR 関数](https://support.office.com/en-us/article/OR-function-7d17ad14-8700-4281-b308-00b131e22af0)| FunctionResult |いずれかの引数が TRUE のときに TRUE を返します。|
|[PDURATION 関数](https://support.office.com/en-us/article/PDURATION-function-44f33460-5be5-4c90-b857-22308892adaf)| FunctionResult |投資が指定した価値に達するまでの投資期間を返します。|
|[PERCENTILE.EXC 関数](https://support.office.com/en-us/article/PERCENTILEEXC-function-bbaa7204-e9e1-4010-85bf-c31dc5dce4ba)| FunctionResult |特定の範囲に含まれるデータの第 k 百分位数に当たる値を返します (k は 0 より大きく 1 より小さい値)。|
|[PERCENTILE.INC 関数](https://support.office.com/en-us/article/PERCENTILEINC-function-680f9539-45eb-410b-9a5e-c1355e5fe2ed)| FunctionResult |特定の範囲に含まれるデータの第 k 百分位数に当たる値を返します。|
|[PERCENTRANK.EXC 関数](https://support.office.com/en-us/article/PERCENTRANKEXC-function-d8afee96-b7e2-4a2f-8c01-8fcdedaa6314)| FunctionResult |データ セット内での値の順位を百分率 (0 より大きく 1 より小さい) で表した値を返します。|
|[PERCENTRANK.INC 関数](https://support.office.com/en-us/article/PERCENTRANKINC-function-149592c9-00c0-49ba-86c1-c1f45b80463a)| FunctionResult |データ セット内での値の順位を百分率で表した値を返します。|
|[PERMUT 関数](https://support.office.com/en-us/article/PERMUT-function-3bd1cb9a-2880-41ab-a197-f246a7a602d3)| FunctionResult |指定された個数のオブジェクトを選択するときの順列の数を返します。|
|[PERMUTATIONA 関数](https://support.office.com/en-us/article/PERMUTATIONA-function-6c7d7fdc-d657-44e6-aa19-2857b25cae4e)| FunctionResult |すべてのオブジェクトから指定された数のオブジェクト (繰り返しを含む) を選択する場合の順列の数を返します。|
|[PHI 関数](https://support.office.com/en-us/article/PHI-function-23e49bc6-a8e8-402d-98d3-9ded87f6295c)| FunctionResult |標準正規分布の密度関数の値を返します。|
|[PI 関数](https://support.office.com/en-us/article/PI-function-264199d0-a3ba-46b8-975a-c4a04608989b)| FunctionResult |円周率 π を返します。|
|[PMT 関数](https://support.office.com/en-us/article/PMT-function-0214da64-9a63-4996-bc20-214433fa6441)| FunctionResult |年間の定期支払額を算出します。|
|[POISSON.DIST 関数](https://support.office.com/en-us/article/POISSONDIST-function-8fe148ff-39a2-46cb-abf3-7772695d9636)| FunctionResult |ポワソン分布の値を返します。|
|[POWER 関数](https://support.office.com/en-us/article/POWER-function-d3f2908b-56f4-4c3f-895a-07fb519c362a)| FunctionResult |数値のべき乗を返します。|
|[PPMT 関数](https://support.office.com/en-us/article/PPMT-function-c370d9e3-7749-4ca4-beea-b06c6ac95e1b)| FunctionResult |指定した期に支払われる投資元金を返します。|
|[PRICE 関数](https://support.office.com/en-us/article/PRICE-function-3ea9deac-8dfa-436f-a7c8-17ea02c21b0a)| FunctionResult |定期的に利息が支払われる証券に対して、額面 $100 あたりの価格を返します。|
|[PRICEDISC 関数](https://support.office.com/en-us/article/PRICEDISC-function-d06ad7c1-380e-4be7-9fd9-75e3079acfd3)| FunctionResult |割引証券の額面 $100 あたりの価格を返します。|
|[PRICEMAT 関数](https://support.office.com/en-us/article/PRICEMAT-function-52c3b4da-bc7e-476a-989f-a95f675cae77)| FunctionResult |満期日に利息が支払われる証券に対して、額面 $100 あたりの価格を返します。|
|[PROB 関数](https://support.office.com/en-us/article/PROB-function-9ac30561-c81c-4259-8253-34f0a238fc49)| FunctionResult |指定した範囲に含まれる値が上限と下限との間に収まる確率を返します。|
|[PRODUCT 関数](https://support.office.com/en-us/article/PRODUCT-function-8e6b5b24-90ee-4650-aeec-80982a0512ce)| FunctionResult |引数を乗算します。|
|[PROPER 関数](https://support.office.com/en-us/article/PROPER-function-52a5a283-e8b2-49be-8506-b2887b889f94)| FunctionResult |文字列に含まれる英単語の先頭文字だけを大文字に変換します。|
|[PV 関数](https://support.office.com/en-us/article/PV-function-23879d31-0e02-4321-be01-da16e8168cbd)| FunctionResult |投資の現在価値を返します。|
|[QUARTILE.EXC 関数](https://support.office.com/en-us/article/QUARTILEEXC-function-5a355b7a-840b-4a01-b0f1-f538c2864cad)| FunctionResult |0 より大きく 1 より小さい百分位値に基づいて、データ セットに含まれるデータから四分位数を返します。|
|[QUARTILE.INC 関数](https://support.office.com/en-us/article/QUARTILEINC-function-1bbacc80-5075-42f1-aed6-47d735c4819d)| FunctionResult |データ セットの四分位数を返します。|
|[QUOTIENT 関数](https://support.office.com/en-us/article/QUOTIENT-function-9f7bf099-2a18-4282-8fa4-65290cc99dee)| FunctionResult |除算の商の整数部を返します。|
|[RADIANS 関数](https://support.office.com/en-us/article/RADIANS-function-ac409508-3d48-45f5-ac02-1497c92de5bf)| FunctionResult |度をラジアンに変換します。|
|[RAND 関数](https://support.office.com/en-us/article/RAND-function-4cbfa695-8869-4788-8d90-021ea9f5be73)| FunctionResult |0 から 1 の乱数を返します。|
|[RANDBETWEEN 関数](https://support.office.com/en-us/article/RANDBETWEEN-function-4cc7f0d1-87dc-4eb7-987f-a469ab381685)| FunctionResult |指定された範囲内の数値の乱数を返します。|
|[RANK.AVG 関数](https://support.office.com/en-us/article/RANKAVG-function-bd406a6f-eb38-4d73-aa8e-6d1c3c72e83a)| FunctionResult |数値のリストの中で、指定した数値の順位を返します。|
|[RANK.EQ 関数](https://support.office.com/en-us/article/RANKEQ-function-284858ce-8ef6-450e-b662-26245be04a40)| FunctionResult |数値のリストの中で、指定した数値の順位を返します。|
|[RATE 関数](https://support.office.com/en-us/article/RATE-function-9f665657-4a7e-4bb7-a030-83fc59e748ce)| FunctionResult |年間の投資金利を返します。|
|[RECEIVED 関数](https://support.office.com/en-us/article/RECEIVED-function-7a3f8b93-6611-4f81-8576-828312c9b5e5)| FunctionResult |全額投資された証券に対して、満期日に支払われる金額を返します。|
|[REPLACE 関数、REPLACEB 関数](https://support.office.com/en-us/article/REPLACE-REPLACEB-functions-8d799074-2425-4a8a-84bc-82472868878a)| FunctionResult |テキスト内の文字を置き換えます。|
|[REPT 関数](https://support.office.com/en-us/article/REPT-function-04c4d778-e712-43b4-9c15-d656582bb061)| FunctionResult |テキストを指定した回数だけ繰り返します。|
|[RIGHT 関数、RIGHTB 関数](https://support.office.com/en-us/article/RIGHT-RIGHTB-functions-240267ee-9afa-4639-a02b-f19e1786cf2f)| FunctionResult |文字列の末尾 (右端) から指定された文字数の文字を返します。|
|[ROMAN 関数](https://support.office.com/en-us/article/ROMAN-function-d6b0b99e-de46-4704-a518-b45a0f8b56f5)| FunctionResult |アラビア数字を、ローマ数字を表す文字列に変換します。|
|[ROUND 関数](https://support.office.com/en-us/article/ROUND-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c)| FunctionResult |数値を四捨五入して指定された桁数にします。|
|[ROUNDDOWN 関数](https://support.office.com/en-us/article/ROUNDDOWN-function-2ec94c73-241f-4b01-8c6f-17e6d7968f53)| FunctionResult |数値を指定された桁数で切り捨てます。|
|[ROUNDUP 関数](https://support.office.com/en-us/article/ROUNDUP-function-f8bc9b23-e795-47db-8703-db171d0c42a7)| FunctionResult |数値を指定された桁数で切り上げます。|
|[ROWS 関数](https://support.office.com/en-us/article/ROWS-function-b592593e-3fc2-47f2-bec1-bda493811597)| FunctionResult |指定の範囲に含まれる行数を返します。|
|[RRI 関数](https://support.office.com/en-us/article/RRI-function-6f5822d8-7ef1-4233-944c-79e8172930f4)| FunctionResult |投資の成長に対する等価利率を返します。|
|[RTD 関数](https://support.office.com/en-us/article/RTD-function-e0cc001a-56f0-470a-9b19-9455dc0eb593)| FunctionResult |COM オートメーションに対応するプログラムからリアルタイムのデータを取得します。|
|[SEC 関数](https://support.office.com/en-us/article/SEC-function-ff224717-9c87-4170-9b58-d069ced6d5f7)| FunctionResult |角度の正割 (セカント) を返します。|
|[SECH 関数](https://support.office.com/en-us/article/SECH-function-e05a789f-5ff7-4d7f-984a-5edb9b09556f)| FunctionResult |角度の双曲線正割を返します。|
|[SECOND 関数](https://support.office.com/en-us/article/SECOND-function-740d1cfc-553c-4099-b668-80eaa24e8af1)| FunctionResult |シリアル値を秒に変換します。|
|[SERIESSUM 関数](https://support.office.com/en-us/article/SERIESSUM-function-a3ab25b5-1093-4f5b-b084-96c49087f637)| FunctionResult |数式で定義されるべき級数の和を返します。|
|[SHEET 関数](https://support.office.com/en-us/article/SHEET-function-44718b6f-8b87-47a1-a9d6-b701c06cff24)| FunctionResult |参照先のシートのシート番号を返します。|
|[SHEETS 関数](https://support.office.com/en-us/article/SHEETS-function-770515eb-e1e8-45ce-8066-b557e5e4b80b)| FunctionResult |参照内のシート数を返します|
|[SIGN 関数](https://support.office.com/en-us/article/SIGN-function-109c932d-fcdc-4023-91f1-2dd0e916a1d8)| FunctionResult |数値の符号を返します。|
|[SIN 関数](https://support.office.com/en-us/article/SIN-function-cf0e3432-8b9e-483c-bc55-a76651c95602)| FunctionResult |指定された角度のサインを返します。|
|[SINH 関数](https://support.office.com/en-us/article/SINH-function-1e4e8b9f-2b65-43fc-ab8a-0a37f4081fa7)| FunctionResult |数値の双曲線正弦を返します。|
|[SKEW 関数](https://support.office.com/en-us/article/SKEW-function-bdf49d86-b1ef-4804-a046-28eaea69c9fa)| FunctionResult |分布の歪度を返します。|
|[SKEW.P 関数](https://support.office.com/en-us/article/SKEWP-function-76530a5c-99b9-48a1-8392-26632d542fcb)| FunctionResult |母集団に基づく分布の歪度を取得します。歪度とは、分布の平均値周辺での両側の非対称度を表す値です。|
|[SLN 関数](https://support.office.com/en-us/article/SLN-function-cdb666e5-c1c6-40a7-806a-e695edc2f1c8)| FunctionResult |定額法 (Straight-line Method) を使用して、資産の 1 期あたりの減価償却費を返します。|
|[SMALL 関数](https://support.office.com/en-us/article/SMALL-function-17da8222-7c82-42b2-961b-14c45384df07)| FunctionResult |指定されたデータ セットの中で k 番目に小さなデータを返します。|
|[SQRT 関数](https://support.office.com/en-us/article/SQRT-function-654975c2-05c4-4831-9a24-2c65e4040fdf)| FunctionResult |正の平方根を返します。|
|[SQRTPI 関数](https://support.office.com/en-us/article/SQRTPI-function-1fb4e63f-9b51-46d6-ad68-b3e7a8b519b4)| FunctionResult |(数値 * π) の平方根を返します。|
|[STANDARDIZE 関数](https://support.office.com/en-us/article/STANDARDIZE-function-81d66554-2d54-40ec-ba83-6437108ee775)| FunctionResult |正規化された値を返します。|
|[STDEV.P 関数](https://support.office.com/en-us/article/STDEVP-function-6e917c05-31a0-496f-ade7-4f4e7462f285)| FunctionResult |母集団全体に基づいて、標準偏差を計算します。|
|[STDEV.S 関数](https://support.office.com/en-us/article/STDEVS-function-7d69cf97-0c1f-4acf-be27-f3e83904cc23)| FunctionResult |標本に基づく標準偏差の推定値を返します。|
|[STDEVA 関数](https://support.office.com/en-us/article/STDEVA-function-5ff38888-7ea5-48de-9a6d-11ed73b29e9d)| FunctionResult |数値、文字列、および論理値を含む標本に基づいて、標準偏差の推定値を返します。|
|[STDEVPA 関数](https://support.office.com/en-us/article/STDEVPA-function-5578d4d6-455a-4308-9991-d405afe2c28c)| FunctionResult |数値、文字列、および論理値を含む母集団全体に基づいて、標準偏差を計算します。|
|[SUBSTITUTE 関数](https://support.office.com/en-us/article/SUBSTITUTE-function-6434944e-a904-4336-a9b0-1e58df3bc332)| FunctionResult |文字列中の指定された文字を他の新しい文字に置き換えます。|
|[SUBTOTAL 関数](https://support.office.com/en-us/article/SUBTOTAL-function-7b027003-f060-4ade-9040-e478765b9939)| FunctionResult |リストまたはデータベースの集計値を返します。|
|[SUM 関数](https://support.office.com/en-us/article/SUM-function-043e1c7d-7726-4e80-8f32-07b23e057f89)| FunctionResult |引数を合計します。|
|[SUMIF 関数](https://support.office.com/en-us/article/SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b)| FunctionResult |指定された検索条件に一致するセルの値を合計します。|
|[SUMIFS 関数](https://support.office.com/en-us/article/SUMIFS-function-c9e748f5-7ea7-455d-9406-611cebce642b)| FunctionResult |セル範囲内で、複数の検索条件を満たすセルの値を合計します。|
|[SUMSQ 関数](https://support.office.com/en-us/article/SUMSQ-function-e3313c02-51cc-4963-aae6-31442d9ec307)| FunctionResult |引数の 2 乗の和 (平方和) を返します。|
|[SYD 関数](https://support.office.com/en-us/article/SYD-function-069f8106-b60b-4ca2-98e0-2a0f206bdb27)| FunctionResult |級数法 (Sum-of-Year's Digits Method) を使用して、特定の期における減価償却費を返します。|
|[T 関数](https://support.office.com/en-us/article/T-function-fb83aeec-45e7-4924-af95-53e073541228)| FunctionResult |引数をテキストに変換します。|
|[T.DIST 関数](https://support.office.com/en-us/article/TDIST-function-4329459f-ae91-48c2-bba8-1ead1c6c21b2)| FunctionResult |スチューデントの t 分布のパーセンテージ (確率) を返します。|
|[T.DIST.2T 関数](https://support.office.com/en-us/article/TDIST2T-function-198e9340-e360-4230-bd21-f52f22ff5c28)| FunctionResult |スチューデントの t 分布のパーセンテージ (確率) を返します。|
|[T.DIST.RT 関数](https://support.office.com/en-us/article/TDISTRT-function-20a30020-86f9-4b35-af1f-7ef6ae683eda)| FunctionResult |スチューデントの t 分布の値を返します。|
|[T.INV 関数](https://support.office.com/en-us/article/TINV-function-2908272b-4e61-4942-9df9-a25fec9b0e2e)| FunctionResult |スチューデントの t 分布の t 値を、確率の関数と自由度で返します。|
|[T.INV.2T 関数](https://support.office.com/en-us/article/TINV2T-function-ce72ea19-ec6c-4be7-bed2-b9baf2264f17)| FunctionResult |スチューデントの t 分布の逆関数値を返します。|
|[TAN 関数](https://support.office.com/en-us/article/TAN-function-08851a40-179f-4052-b789-d7f699447401)| FunctionResult |数値の正接 (タンジェント) を返します。|
|[TANH 関数](https://support.office.com/en-us/article/TANH-function-017222f0-a0c3-4f69-9787-b3202295dc6c)| FunctionResult |数値の双曲線正接を返します。|
|[TBILLEQ 関数](https://support.office.com/en-us/article/TBILLEQ-function-2ab72d90-9b4d-4efe-9fc2-0f81f2c19c8c)| FunctionResult |米国財務省短期証券 (TB) の債券換算利回りを返します。|
|[TBILLPRICE 関数](https://support.office.com/en-us/article/TBILLPRICE-function-eacca992-c29d-425a-9eb8-0513fe6035a2)| FunctionResult |米国財務省短期証券 (TB) の額面 $100 あたりの価格を返します。|
|[TBILLYIELD 関数](https://support.office.com/en-us/article/TBILLYIELD-function-6d381232-f4b0-4cd5-8e97-45b9c03468ba)| FunctionResult |米国財務省短期証券 (TB) の利回りを返します。|
|[TEXT 関数](https://support.office.com/en-us/article/TEXT-function-20d5ac4d-7b94-49fd-bb38-93d29371225c)| FunctionResult |数値を、書式設定したテキストに変換します。|
|[TIME 関数](https://support.office.com/en-us/article/TIME-function-9a5aff99-8f7d-4611-845e-747d0b8d5457)| FunctionResult |指定した時刻に対応するシリアル値を返します。|
|[TIMEVALUE 関数](https://support.office.com/en-us/article/TIMEVALUE-function-0b615c12-33d8-4431-bf3d-f3eb6d186645)| FunctionResult |時刻を表す文字列をシリアル値に変換します。|
|[TODAY 関数](https://support.office.com/en-us/article/TODAY-function-5eb3078d-a82c-4736-8930-2f51a028fdd9)| FunctionResult |現在の日付に対応するシリアル値を返します。|
|[TRIM 関数](https://support.office.com/en-us/article/TRIM-function-410388fa-c5df-49c6-b16c-9e5630b479f9)| FunctionResult |テキストからスペースを削除します。|
|[TRIMMEAN 関数](https://support.office.com/en-us/article/TRIMMEAN-function-d90c9878-a119-4746-88fa-63d988f511d3)| FunctionResult |データ セットの中間項の平均を返します。|
|[TRUE 関数](https://support.office.com/en-us/article/TRUE-function-7652c6e3-8987-48d0-97cd-ef223246b3fb)| FunctionResult |論理値 TRUE を返します。|
|[TRUNC 関数](https://support.office.com/en-us/article/TRUNC-function-8b86a64c-3127-43db-ba14-aa5ceb292721)| FunctionResult |数値の小数部を切り捨てて整数にします。|
|[TYPE 関数](https://support.office.com/en-us/article/TYPE-function-45b4e688-4bc3-48b3-a105-ffa892995899)| FunctionResult |値のデータ型を表す数値を返します。|
|[UNICHAR 関数](https://support.office.com/en-us/article/UNICHAR-function-ffeb64f5-f131-44c6-b332-5cd72f0659b8)| FunctionResult |指定された数値により参照される Unicode 文字を返します。|
|[UNICODE 関数](https://support.office.com/en-us/article/UNICODE-function-adb74aaa-a2a5-4dde-aff6-966e4e81f16f)| FunctionResult |文字列の最初の文字に対応する番号 (コード ポイント) を返します。|
|[UPPER 関数](https://support.office.com/en-us/article/UPPER-function-c11f29b3-d1a3-4537-8df6-04d0049963d6)| FunctionResult |文字列に含まれる英字をすべて大文字に変換します。|
|[VALUE 関数](https://support.office.com/en-us/article/VALUE-function-257d0108-07dc-437d-ae1c-bc2d3953d8c2)| FunctionResult |テキスト引数を数値に変換します。|
|[VAR.P 関数](https://support.office.com/en-us/article/VARP-function-73d1285c-108c-4843-ba5d-a51f90656f3a)| FunctionResult |母集団全体に基づいて、分散を計算します。|
|[VAR.S 関数](https://support.office.com/en-us/article/VARS-function-913633de-136b-449d-813e-65a00b2b990b)| FunctionResult |標本に基づいて、分散の推定値を返します。|
|[VARA 関数](https://support.office.com/en-us/article/VARA-function-3de77469-fa3a-47b4-85fd-81758a1e1d07)| FunctionResult |数値、文字列、および論理値を含む標本に基づいて、分散の推定値を返します。|
|[VARPA 関数](https://support.office.com/en-us/article/VARPA-function-59a62635-4e89-4fad-88ac-ce4dc0513b96)| FunctionResult |数値、文字列、および論理値を含む母集団全体に基づいて、分散を計算します。|
|[VDB 関数](https://support.office.com/en-us/article/VDB-function-dde4e207-f3fa-488d-91d2-66d55e861d73)| FunctionResult |定率法 (declining Balance Method) を利用して、特定の期または部分的な期における資産の減価償却費を返します。|
|[VLOOKUP 関数](https://support.office.com/en-us/article/VLOOKUP-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1)| FunctionResult |配列の左端列で特定の値を検索し、その行内で移動して、対応するセルの値を返します。|
|[WEEKDAY 関数](https://support.office.com/en-us/article/WEEKDAY-function-60e44483-2ed1-439f-8bd0-e404c190949a)| FunctionResult |シリアル値を曜日に変換します。|
|[WEEKNUM 関数](https://support.office.com/en-us/article/WEEKNUM-function-e5c43a03-b4ab-426c-b411-b18c13c75340)| FunctionResult |シリアル値をその年の何週目に当たるかを示す値に変換します。|
|[WEIBULL.DIST 関数](https://support.office.com/en-us/article/WEIBULLDIST-function-4e783c39-9325-49be-bbc9-a83ef82b45db)| FunctionResult |ワイブル分布の値を返します。|
|[WORKDAY 関数](https://support.office.com/en-us/article/WORKDAY-function-f764a5b7-05fc-4494-9486-60d494efbf33)| FunctionResult |指定した稼動日数だけ前または後の日付に対応するシリアル値を返します。|
|[WORKDAY.INTL 関数](https://support.office.com/en-us/article/WORKDAYINTL-function-a378391c-9ba7-4678-8a39-39611a9bf81d)| FunctionResult |週末がどの曜日で何日間あるかを示すパラメーターを使用して、指定した稼働日数だけ前または後の日付に対応するシリアル値を返します。|
|[XIRR 関数](https://support.office.com/en-us/article/XIRR-function-de1242ec-6477-445b-b11b-a303ad9adc9d)| FunctionResult |定期的でないキャッシュ フローの特定のスケジュールに対する内部利益率を返します。|
|[XNPV 関数](https://support.office.com/en-us/article/XNPV-function-1b42bbf6-370f-4532-a0eb-d67c16b664b7)| FunctionResult |定期的でないキャッシュ フローの特定のスケジュールに対する正味現在価値を返します。|
|[XOR 関数](https://support.office.com/en-us/article/XOR-function-1548d4c2-5e47-4f77-9a92-0533bba14f37)| FunctionResult |すべての引数の論理排他 OR を返します。|
|[YEAR 関数](https://support.office.com/en-us/article/YEAR-function-c64f017a-1354-490d-981f-578e8ec8d3b9)| FunctionResult |シリアル値を年に変換します。|
|[YEARFRAC 関数](https://support.office.com/en-us/article/YEARFRAC-function-3844141e-c76d-4143-82b6-208454ddc6a8)| FunctionResult |開始日と終了日を指定して、その間の期間が 1 年間に対して占める割合を返します。|
|[YIELD 関数](https://support.office.com/en-us/article/YIELD-function-f5f5ca43-c4bd-434f-8bd2-ed3c9727a4fe)| FunctionResult |利息が定期的に支払われる証券の利回りを返します。|
|[YIELDDISC 関数](https://support.office.com/en-us/article/YIELDDISC-function-a9dbdbae-7dae-46de-b995-615faffaaed7)| FunctionResult |米国財務省短期証券 (TB) などの割引債の年利回りを返します。|
|[YIELDMAT 関数](https://support.office.com/en-us/article/YIELDMAT-function-ba7d1809-0d33-4bcb-96c7-6c56ec62ef6f)| FunctionResult |満期日に利息が支払われる証券の利回りを返します。|
|[Z.TEST 関数](https://support.office.com/en-us/article/ZTEST-function-d633d5a3-2031-4614-a016-92180ad82bee)| FunctionResult |Z 検定の片側確率の値を返します。|

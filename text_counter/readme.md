# text_counter

特定の単語が資料に何回出現しているかをカウントします。
This python code counts how many times spesific keywords appear.

## 使用例
家具や設備などの配置が記入されたpptx書類から、自動で集計が可能。

## 使い方
1. data/inoput.xlsx を開き、"input"と名付けてあるsheetに、カウントしたいキーワード群をセットしてください
2. "url"と名付けてあるsheetに、カウント対象のpptxファイルの絶対パスを入力してください。
3. text_counter.pyを実行
4. 再度 data/inoput.xlsxを開き、"result"と名付けてあるsheetで結果が確認できます。

## 使用にあたっての注意点
図形に隠れていたり、表示領域からはみ出ているテキストもカウント対象となります。

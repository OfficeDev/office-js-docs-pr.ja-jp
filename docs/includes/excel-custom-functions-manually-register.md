オートコンププリート メニューで `CONTOSO` 名前空間を使用できない場合、次の手順でアドインを Excel に登録します。

### <a name="excel-on-windows-or-mac"></a>[Windows または Mac 上の Excel](#tab/excel-windows)

1. Excel で、**[挿入]** タブを選択し、**[個人用アドイン]** の右側に配置された下矢印を選択します。

    :::image type="content" source="../images/select-insert.png" alt-text="[個人用アドイン] の下矢印が強調表示された Windows での Excel の [挿入] リボンのスクリーンショット":::

1. 使用可能なアドインのリストから [**開発者向けアドイン**] セクションを見つけ、**starcount** アドインを選択して登録します。

    :::image type="content" source="../images/list-starcount.png" alt-text="[個人用アドイン] 一覧で [Excel カスタム関数] アドインが強調表示されている Windows 上の Excel の [挿入] リボンのスクリーンショット。":::

# <a name="excel-on-the-web"></a>[Excel on the web](#tab/excel-online)

1. Excel で、[**挿入**] タブを選択して、[**アドイン**] を選択します。

    :::image type="content" source="../images/excel-cf-online-register-add-in-1.png" alt-text="[個人用アドイン] ボタンが強調表示された Excel on the web の [挿入] リボンのスクリーンショット。":::

1. **[マイ アドインの管理]** を選択し、**[マイ アドインのアップロード]** を選択します。

1. **[参照...]** を選択し、Yeoman ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。

1. **manifest.xml** ファイルを選択し、**[開く]** を選択し、**[アップロード]** を選択します。

1. 新しい関数をお試しください。 セル **B1** で、テキスト **=CONTOSO.GETSTARCOUNT("OfficeDev", "Excel-Custom-Functions")** を入力し、Enter キーを押します。 セル **B1** の結果は [Excel-Custom-Functions Github リポジトリ](https://github.com/OfficeDev/Excel-Custom-Functions) に与えられた現在の星の数です。

---

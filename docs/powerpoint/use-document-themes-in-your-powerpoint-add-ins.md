# <a name="use-document-themes-in-your-powerpoint-add-ins"></a>PowerPoint アドインでドキュメントのテーマを使用する


[Office テーマ](https://support.office.com/en-US/Article/What-is-a-theme--7528ccc2-4327-4692-8bf5-9b5a3f2a5ef5)の一部は表示が調整されたフォントと色のセットで構成されており、このセットをプレゼンテーション、文書、ワークシート、メールに適用できます。PowerPoint でプレゼンテーションのテーマを適用したりカスタマイズしたりするには、リボンの **[デザイン]** タブの **[テーマ]** グループと **[バリエーション]** グループを使います。PowerPoint は既定の **Office テーマ**の新しい空白のプレゼンテーションを割り当てますが、**[デザイン]** タブ上の使用できる他のテーマを選択したり、Office.com から追加のテーマをダウンロードしたり、独自のテーマを作成してカスタマイズしたりできます。

OfficeThemes.css を使用すると、次の 2 つの方法でアドインを PowerPoint に合わせて設計できます。


-  **PowerPoint 用のコンテンツ アドイン**。OfficeThemes.css の文書テーマ クラスを使って、コンテンツ アドイン の挿入先のプレゼンテーションのテーマと一致するフォントと色を指定します。このフォントと色は、ユーザーがプレゼンテーションのテーマを変更したりカスタマイズしたりすると動的に更新されます。
    
-  **PowerPoint 用の作業ウィンドウ アドイン**。OfficeThemes.css の Office UI テーマ クラスを使って、この UI で使用されているフォントと背景の色を指定し、作業ウィンドウ アドインが組み込み作業ウィンドウの色と一致するようにします。この色は、ユーザーが Office UI テーマを変更すると動的に更新されます。
    

### <a name="document-theme-colors"></a>文書のテーマの色

すべての Office 文書のテーマには 12 色が定義されています。これらのうち 10 色は、色選択を使ってプレゼンテーション内のフォントや背景などの色を設定するときに利用できます。


![カラー パレット](../../images/off15app_ColorPalette.png)

PowerPoint でテーマの 12 色をすべて表示したりカスタマイズしたりするには、**[デザイン]** タブの **[バリエーション]** グループで、**[その他]** ドロップダウンをクリックしてから、**[色]** をポイントし、**[色のカスタマイズ]** をクリックして **[新しいテーマの色の作成]** ダイアログ ボックスを表示します。


![新しいテーマの色のダイアログ ボックスの作成](../../images/off15app_CreateNewThemeColors.png)

最初の 4 色はテキストと背景用です。テキストを明るい色で作成すると常に暗い色より読みやすくなり、テキストを暗い色で作成すると常に明るい色より読みやすくなります。続く 6 色は、4 つの背景になる色の上に常に表示されるアクセントです。最後の 2 色は、ハイパーリンクと表示済みハイパーリンクの色です。


### <a name="document-theme-fonts"></a>文書のテーマのフォント

すべての Office 文書のテーマには 2 つのフォント (見出し用と本文テキスト用) も定義されています。PowerPoint はこれらのフォントを使って自動的にテキスト スタイルを構成します。また、テキストと**ワードアート**の**クイック スタイル** ギャラリーでも同じテーマのフォントを使います。これらの 2 つのフォントは、フォント ピッカーを使ってフォントを選択するときに、最初の 2 つの選択項目として利用できます。


![フォント ピッカー](../../images/off15app_FontPicker.png)

PowerPoint でテーマを表示したりカスタマイズしたりするには、**[デザイン]** タブの **[バリエーション]** グループで、**[その他]** ドロップダウンをクリックしてから、**[フォント]** をポイントし、**[フォントのカスタマイズ]** をクリックして **[新しいテーマのフォント パターンの作成]** ダイアログ ボックスを表示します。


![新しいテーマのフォントのダイアログ ボックスの作成](../../images/off15app_CreateNewThemeFonts.png)


### <a name="office-ui-theme-fonts-and-colors"></a>Office の UI のテーマのフォントと色

Office では、すべての Office アプリケーションの UI で使用される色とフォントの一部を指定している複数の事前定義済みのテーマの中から選択することもできます。そのためには、[ **ファイル**]  >  [ **アカウント**]  >  [ **Office テーマ**] ドロップダウンを (Office アプリケーションから) 使用します。


![Office テーマ ドロップ ダウン](../../images/off15app_OfficeThemePicker.png)

OfficeThemes.css には PowerPoint 用の作業ウィンドウ アドインで使用できるクラスが含まれており、両者が使用するフォントと色は同じになります。したがって、組み込み作業ウィンドウの外観と一致する作業ウィンドウ アドインを設計できます。


## <a name="using-officethemescss"></a>OfficeThemes.css を使用する

OfficeThemes.css ファイルと PowerPoint 用のコンテンツ アドインを併用すると、アドイン の外観を、一緒に実行するプレゼンテーションに適用されているテーマに合わせて調整できます。OfficeThemes.css ファイルと PowerPoint 用の作業ウィンドウ アドインを併用すると、Office UI のフォントと色に合わせて アドイン を調整できます。


### <a name="adding-the-officethemescss-file-to-your-project"></a>OfficeThemes.css ファイルをプロジェクトに追加する

OfficeThemes.css ファイルを アドイン プロジェクトに追加して、このファイルを参照するには、次の手順に従います。


### <a name="to-add-officethemescss-to-your-visual-studio-project"></a>OfficeThemes.css を Visual Studio プロジェクトに追加するには


1. **ソリューション エクスプローラー**で、****project_name****_Web_ プロジェクト内の **[コンテンツ]** フォルダーを右クリックし、**[追加]** をポイントしてから、**[スタイル シート]** をクリックします。
    
2. 新しいスタイル シートに OfficeThemes という名前を付けます。
    
     >**重要** このスタイル シートは OfficeThemes という名前にする必要があります。ユーザーがテーマを変更すると、アドインのフォントと色を動的に更新する機能は動作しなくなります。
3. ファイル内の既定の  **body** クラス ( `body {}`) を削除し、次の CSS コードをコピーしてファイルに貼り付けます。
    
```
  /* The following classes describe the common theme information for office documents */ /* Basic Font and Background Colors for text */ .office-docTheme-primary-fontColor { color:#000000; } .office-docTheme-primary-bgColor { background-color:#ffffff; } .office-docTheme-secondary-fontColor { color: #000000; } .office-docTheme-secondary-bgColor { background-color: #ffffff; } /* Accent color definitions for fonts */ .office-contentAccent1-color { color:#5b9bd5; } .office-contentAccent2-color { color:#ed7d31; } .office-contentAccent3-color { color:#a5a5a5; } .office-contentAccent4-color { color:#ffc000; } .office-contentAccent5-color { color:#4472c4; } .office-contentAccent6-color { color:#70ad47; } /* Accent color for backgrounds */ .office-contentAccent1-bgColor { background-color:#5b9bd5; } .office-contentAccent2-bgColor { background-color:#ed7d31; } .office-contentAccent3-bgColor { background-color:#a5a5a5; } .office-contentAccent4-bgColor { background-color:#ffc000; } .office-contentAccent5-bgColor { background-color:#4472c4; } .office-contentAccent6-bgColor { background-color:#70ad47; } /* Accent color for borders */ .office-contentAccent1-borderColor { border-color:#5b9bd5; } .office-contentAccent2-borderColor { border-color:#ed7d31; } .office-contentAccent3-borderColor { border-color:#a5a5a5; } .office-contentAccent4-borderColor { border-color:#ffc000; } .office-contentAccent5-borderColor { border-color:#4472c4; } .office-contentAccent6-borderColor { border-color:#70ad47; } /* links */ .office-a { color: #0563c1; } .office-a:visited { color: #954f72; } /* Body Fonts */ .office-bodyFont-eastAsian { } /* East Asian name of the Font */ .office-bodyFont-latin { font-family:"Calibri"; } /* Latin name of the Font */ .office-bodyFont-script { } /* Script name of the Font */ .office-bodyFont-localized { font-family:"Calibri"; } /* Localized name of the Font. Corresponds to the default font of the culture currently used in Office.*/ /* Headers Font */ .office-headerFont-eastAsian { } .office-headerFont-latin { font-family:"Calibri Light"; } .office-headerFont-script { } .office-headerFont-localized { font-family:"Calibri Light"; } /* The following classes define font and background colors for Office UI themes. These classes should only be used in task pane add-ins */ /* Basic Font and Background Colors for PPT */ .office-officeTheme-primary-fontColor { color:#b83b1d; } .office-officeTheme-primary-bgColor { background-color:#dedede; } .office-officeTheme-secondary-fontColor { color:#262626; } .office-officeTheme-secondary-bgColor { background-color:#ffffff; } 
```

4. Visual Studio 以外のツールを使用して アドイン を作成している場合は、手順 3 の CSS コードをテキスト ファイルにコピーし、OfficeThemes.css としてファイルを保存したことを確認してください。
    

### <a name="referencing-officethemescss-in-your-add-ins-html-pages"></a>アドインの HTML ページ内で OfficeThemes.css を参照する

アドイン プロジェクト内で OfficeThemes.css ファイルを使用するには、OfficeThemes.css ファイルを参照する  `<link>` タグを、アドイン の UI を実装する Web ページ (.html, .aspx, .php ファイルなど) の `<head>` タグ内に、以下の形式で追加します。


```HTML
<link href="<local_path_to_OfficeThemes.css> " rel="stylesheet" type="text/css" />
```

Visual Studio でこの作業を行うには、次の手順に従ってください。


### <a name="to-reference-officethemescss-in-your-add-in-for-powerpoint"></a>PowerPoint 用アドイン内で OfficeThemes.css を参照するには


1. Visual Studio 2015 以降で、 **Office アドイン** プロジェクトを開くか新規作成します。
    
2. アドイン の UI を実装する HTML ページ (既定のテンプレート内の Home.html など) で、OfficeThemes.css ファイルを参照する次の  `<link>` タグを `<head>` タグに追加します。
    
```HTML
  <link href="../../Content/OfficeThemes.css" rel="stylesheet" type="text/css" />
```

Visual Studio 以外のツールで アドイン を作成している場合は、アドイン に展開する OfficeThemes.css のコピーへの相対パスを指定して、同じ形式の  `<link>` タグを追加してください。


### <a name="using-officethemescss-document-theme-classes-in-your-content-add-ins-html-page"></a>コンテンツ アドインの HTML ページで OfficeThemes.css 文書テーマ クラスを使用する

OfficeTheme.css 文書テーマ クラスを使用するコンテンツ アドイン 内の HTML の簡単な例を以下に示します。文書のテーマで使用される 12 色と 2 つのフォントに対応する OfficeThemes.css クラスの詳細については、「 [コンテンツ アドインのテーマ クラス](#theme-classes-for-content-add-ins)」を参照してください。


```HTML
<body> <div id="themeSample" class="office-docTheme-primary-fontColor "> <h1 class="office-headerFont-latin">Hello world!</h1> <h1 class="office-headerFont-latin office-contentAccent1-bgColor">Hello world!</h1> <h1 class="office-headerFont-latin office-contentAccent2-bgColor">Hello world!</h1> <h1 class="office-headerFont-latin office-contentAccent3-bgColor">Hello world!</h1> <h1 class="office-headerFont-latin office-contentAccent4-bgColor">Hello world!</h1> <h1 class="office-headerFont-latin office-contentAccent5-bgColor">Hello world!</h1> <h1 class="office-headerFont-latin office-contentAccent6-bgColor">Hello world!</h1> <p class="office-bodyFont-latin office-docTheme-secondary-fontColor">Hello world!</p> </div> </body>
```

実行時に、既定の  **Office テーマ**を使用するプレゼンテーションにコンテンツ アドイン を挿入すると、次のように表示されます。


![Office のテーマを使用して実行しているコンテンツ アプリ](../../images/off15app_ContentApp_OfficeTheme.png)

別のテーマを使うようにプレゼンテーションを変更するか、プレゼンテーションのテーマをカスタマイズすると、OfficeThemes.css クラスで指定されたフォントと色は、プレゼンテーションのテーマのフォントと色に対応するように動的に更新されます。前述の例と同じ HTML を使うと、アドインの挿入先のプレゼンテーションで**ファセット**のテーマが使われて、アドインは次のように表示されます。


![ファセットのテーマを使用して実行しているコンテンツ アプリ](../../images/off15app_ContentApp_FacetTheme.png)


### <a name="using-officethemescss-office-ui-theme-classes-in-your-task-pane-add-ins-html-page"></a>作業ウィンドウ アドインの HTML ページで OfficeThemes.css Office UI テーマ クラスを使用する

ユーザーは、文書のテーマに加えて、すべての Office アプリケーションの Office ユーザー インターフェイスの配色をカスタマイズできます。そのためには、 [ **ファイル**]  >  [ **アカウント**]  >  [ **Office テーマ**] ドロップ ダウン ボックスを使用します。

OfficeTheme.css クラスを使用してフォントの色と背景色を指定する作業ウィンドウ アドイン 内の HTML の簡単な例を次に示します。Office UI テーマのフォントと色に対応する OfficeThemes.css クラスの詳細については、「[作業ウィンドウ アドインのテーマ クラス](#theme-classes-for-task-pane-add-ins)」を参照してください。


```HTML
<body> <div id="content-header" class="office-officeTheme-primary-fontColor office-officeTheme-primary-bgColor"> <div class="padding"> <h1>Welcome</h1> </div> </div> <div id="content-main" class="office-officeTheme-secondary-fontColor office-officeTheme-secondary-bgColor"> <div class="padding"> <p>Add home screen content here.</p> <p>For example:</p> <button id="get-data-from-selection">Get data from selection</button> <p> <a target="_blank" class="office-a" href="https://go.microsoft.com/fwlink/?LinkId=276812"> Find more samples online... </a> </p> </div> </div> </body> 
```

PowerPoint で [ **ファイル**]  >  [ **アカウント**]  >  [ **Office テーマ**] を [ **白**] に設定して実行すると、作業ウィンドウ アドイン は次のように表示されます。


![白の Office の作業ウィンドウ](../../images/off15app_TaskPaneThemeWhite.png)

**[Office テーマ]** を **[濃い灰色]** に変更すると、OfficeThemes.css クラスで指定されたフォントと色は動的に更新されて次のように表示されます。


![濃い灰色の Office の作業ウィンドウ](../../images/off15app_TaskPaneThemeDarkGray.png)


## <a name="officethemecss-classes"></a>OfficeTheme.css のクラス


OfficeThemes.css には、PowerPoint 用のコンテンツ アドインおよび作業ウィンドウ アドインと併用できる 2 つのクラスのセットが含まれています。


### <a name="theme-classes-for-content-add-ins"></a>コンテンツ アドインのテーマ クラス


OfficeThemes.css ファイルには、文書テーマで使用される 2 つのフォントと 12 色に対応するクラスがあります。これらのクラスの適切な使用法は、PowerPoint 用のコンテンツ アドインと併用して、アドインのフォントと色が挿入先のプレゼンテーションに合わせて調整されるようにすることです。


**コンテンツ アドインのテーマ フォント**


|**クラス**|**説明**|
|:-----|:-----|
| `office-bodyFont-eastAsian`|本文のフォントの東アジア言語の名前。|
| `office-bodyFont-latin`|本文のフォントのラテン文字の名前。既定は「Calabri」です。|
| `office-bodyFont-script`|本文のフォントのスクリプト名。|
| `office-bodyFont-localized`|本文のフォントのローカライズされた名前。Office で現在使用されているカルチャに従って既定のフォント名を指定します。|
| `office-headerFont-eastAsian`|ヘッダーのフォントの東アジア言語の名前。|
| `office-headerFont-latin`|ヘッダーのフォントのラテン文字の名前。既定は「Calabri Light」です。|
| `office-headerFont-script`|ヘッダーのフォントのスクリプト名。|
| `office-headerFont-localized`|ヘッダーのフォントのローカライズされた名。Office で現在使用されているカルチャに従って既定のフォント名を指定します。|

**コンテンツ アドインのテーマの色**


|**クラス**|**説明**|
|:-----|:-----|
| `office-docTheme-primary-fontColor`|第 1 フォントの色。既定は #000000 です。|
| `office-docTheme-primary-bgColor`|第 1 フォントの背景色。既定は #FFFFFF です。|
| `office-docTheme-secondary-fontColor`|第 2 フォントの色。既定は #000000 です。|
| `office-docTheme-secondary-bgColor`|第 2 フォントの背景色。既定は #FFFFFF です。|
| `office-contentAccent1-color`|フォントのアクセント 1。既定は #5B9BD5 です。|
| `office-contentAccent2-color`|フォントのアクセント 2。既定は #ED7D31 です。|
| `office-contentAccent3-color`|フォントのアクセント 3。既定は #A5A5A5 です。|
| `office-contentAccent4-color`|フォントのアクセント 4。既定は #FFC000 です。|
| `office-contentAccent5-color`|フォントのアクセント 5。既定は #4472C4 です。|
| `office-contentAccent6-color`|フォントのアクセント 6。既定は #70AD47 です。|
| `office-contentAccent1-bgColor`|背景のアクセント 1。既定は #5B9BD5 です。|
| `office-contentAccent2-bgColor`|背景のアクセント 2。既定は #ED7D31 です。|
| `office-contentAccent3-bgColor`|背景のアクセント 3。既定は #A5A5A5 です。|
| `office-contentAccent4-bgColor`|背景のアクセント 4。既定は #FFC000 です。|
| `office-contentAccent5-bgColor`|背景のアクセント 5。既定は #4472C4 です。|
| `office-contentAccent6-bgColor`|背景のアクセント 6。既定は #70AD47 です。|
| `office-contentAccent1-borderColor`|境界線のアクセント 1。既定は #5B9BD5 です。|
| `office-contentAccent2-borderColor`|境界線のアクセント 2。既定は #ED7D31 です。|
| `office-contentAccent3-borderColor`|境界線のアクセント 3。既定は #A5A5A5 です。|
| `office-contentAccent4-borderColor`|境界線のアクセント 4。既定は #FFC000 です。|
| `office-contentAccent5-borderColor`|境界線のアクセント 5。既定は #4472C4 です。|
| `office-contentAccent6-borderColor`|境界線のアクセント 6。既定は #70AD47 です。|
| `office-a`|ハイパーリンクの色。既定は #0563C1 です。|
| `office-a:visited`|表示済みのハイパーリンクの色。既定は #954F72 です。|
次のスクリーンショットは、既定の Office テーマの使用時に アドイン テキストに割り当てられるテーマの色のクラスすべて (2 つのハイパーリンクの色を除く) の例を示しています。


![既定の Office テーマの色の例](../../images/off15app_DefaultOfficeThemeColors.png)


### <a name="theme-classes-for-task-pane-add-ins"></a>作業ウィンドウ アドインのテーマ クラス


OfficeThemes.css ファイルには、Office アプリケーション UI テーマで使用されるフォントと背景に割り当てられた 4 色に対応するクラスがあります。これらのクラスの適切な使用法は、PowerPoint 用の作業ウィンドウ アドインと併用して、アドインの色が Office 内の他の組み込み作業ウィンドウに合わせて調整されるようにすることです。


**作業ウィンドウ アドインのテーマのフォントと背景色**


|**クラス**|**説明**|
|:-----|:-----|
| `office-officeTheme-primary-fontColor`|第 1 フォントの色。既定: # B83B1D|
| `office-officeTheme-primary-bgColor`|第 1 背景色。既定は #DEDEDE です。|
| `office-officeTheme-secondary-fontColor`|第 2 フォントの色。既定は 262626 です。|
| `office-officeTheme-secondary-bgColor`|第 2 背景色。既定は #FFFFFF です。|

## <a name="additional-resources"></a>その他のリソース

- [PowerPoint 用のコンテンツ アドインと作業ウィンドウ アドインを作成する](../powerpoint/powerpoint-add-ins.md)

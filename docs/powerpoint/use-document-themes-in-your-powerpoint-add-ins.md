---
title: PowerPoint アドインでドキュメントのテーマを使用する
description: プレゼンテーション、ドキュメント、ワークシート、電子メールに適用するフォントや色などのテーマを視覚的に調整する方法について説明します。
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: b3dfce5e36f1247e84cdd13372fd822a27a59fe0
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938382"
---
# <a name="use-document-themes-in-your-powerpoint-add-ins"></a>PowerPoint アドインでドキュメントのテーマを使用する

[Office テーマ](https://support.microsoft.com/office/83e68627-2c17-454a-9fd8-62deb81951a6)の一部は表示が調整されたフォントと色のセットで構成されており、このセットをプレゼンテーション、文書、ワークシート、電子メールに適用できます。PowerPoint でプレゼンテーションのテーマを適用したりカスタマイズしたりするには、リボンの **[デザイン]** タブの **[テーマ]** グループと **[バリエーション]** グループを使います。PowerPoint は既定の **Office テーマ** の新しい空白のプレゼンテーションを割り当てますが、**[デザイン]** タブ上の使用できる他のテーマを選択したり、Office.com から追加のテーマをダウンロードしたり、独自のテーマを作成してカスタマイズしたりできます。

OfficeThemes.css を使用して、2 つの方法でPowerPointアドインを設計します。

- **PowerPoint 用のコンテンツ アドイン**。OfficeThemes.css の文書テーマ クラスを使用して、コンテンツ アドインの挿入先のプレゼンテーションのテーマと一致するフォントと色を指定します。このフォントと色は、ユーザーがプレゼンテーションのテーマを変更したりカスタマイズしたりすると動的に更新されます。
    
- **PowerPoint 用の作業ウィンドウ アドイン**。OfficeThemes.css の Office UI テーマ クラスを使用して、この UI で使用されているフォントと背景の色を指定し、作業ウィンドウ アドインが組み込み作業ウィンドウの色と一致するようにします。この色は、ユーザーが Office UI テーマを変更すると動的に更新されます。

## <a name="document-theme-colors"></a>文書のテーマの色

すべての Office 文書のテーマには 12 色が定義されています。これらのうち 10 色は、色選択を使ってプレゼンテーション内のフォントや背景などの色を設定するときに利用できます。

![カラー パレット。](../images/office15-app-color-palette.png)

PowerPoint で 12 種類のテーマの色の完全なセットを表示またはカスタマイズするには、[デザイン]タブの [バリアント] グループで、[その他] ドロップダウンをクリックし、[色のカスタマイズ] を選択して [新しいテーマの色の作成] ダイアログ ボックスを表示します。   >   

![[新しいテーマの色を作成する] ダイアログ ボックス。](../images/office15-app-create-new-theme-colors.png)

最初の 4 色はテキストと背景用です。テキストを明るい色で作成すると常に暗い色より読みやすくなり、テキストを暗い色で作成すると常に明るい色より読みやすくなります。続く 6 色は、4 つの背景になる色の上に常に表示されるアクセントです。最後の 2 色は、ハイパーリンクと表示済みハイパーリンクの色です。

## <a name="document-theme-fonts"></a>文書のテーマのフォント

すべての Office 文書のテーマには 2 つのフォント (見出し用と本文テキスト用) も定義されています。 PowerPoint はこれらのフォントを使用して自動的にテキスト スタイルを構成します。 また、テキストと **ワードアート** の **クイック スタイル** ギャラリーでも同じフォントが使用されます。 これらの 2 つのフォントは、フォント ピッカーを使用してフォントを選択するときに、最初の 2 つの選択項目として利用できます。

![フォントピッカー。](../images/office15-app-font-picker.png)

PowerPoint でテーマ フォントを表示またはカスタマイズするには、[デザイン]タブの [バリアント] グループで、[その他] ドロップダウンをクリックし、[フォントのカスタマイズ] を選択して [新しいテーマ フォントの作成] ダイアログ ボックスを表示します。   >   

![[新しいテーマ フォントの作成] ダイアログ ボックス。](../images/office15-app-create-new-theme-fonts.png)

### <a name="office-ui-theme-fonts-and-colors"></a>Office の UI のテーマのフォントと色

Office also lets you choose between several predefined themes that specify some of the colors and fonts used in the UI of all Office applications. これを行うには、[ファイル アカウント]**ドロップダウンOffice**(任意のアプリケーションから)  >    >   Officeします。

![Officeのドロップダウン。](../images/office15-app-office-theme-picker.png)

OfficeThemes.css には PowerPoint 用の作業ウィンドウ アドインで使用できるクラスが含まれており、両者が使用するフォントと色は同じになります。したがって、組み込み作業ウィンドウの外観と一致する作業ウィンドウ アドインを設計できます。

## <a name="use-officethemescss"></a>OfficeThemes.css を使用する

OfficeThemes.css ファイルと PowerPoint 用のコンテンツ アドインを併用すると、アドイン の外観を、一緒に実行するプレゼンテーションに適用されているテーマに合わせて調整できます。OfficeThemes.css ファイルと PowerPoint 用の作業ウィンドウ アドインを併用すると、Office UI のフォントと色に合わせて アドイン を調整できます。

## <a name="add-the-officethemescss-file-to-your-project"></a>OfficeThemes.css ファイルをプロジェクトに追加する

OfficeThemes.css ファイルを アドイン プロジェクトに追加して、このファイルを参照するには、次の手順に従います。

> [!NOTE]
> この手順は、Visual Studio 2015 にのみ適用されます。 2019 Visual Studioを使用している場合、OfficeThemes.css ファイルは、作成する新しい PowerPoint アドイン プロジェクトに対して自動的に作成されます。

1. **ソリューション エクスプローラー** で、_**project_name**_**Web** プロジェクト内の [**コンテンツ**] フォルダーを右クリックし、[**追加**] を選択してから、[**スタイル シート**] を選択します。
    
2. 新しいスタイル シートに **OfficeThemes** という名前を付けます。
    
   > [!IMPORTANT]
   > スタイル シートの名前は OfficeThemes と指定する必要があります。そうでない場合、ユーザーがテーマを変更した際にアドインのフォントと色を動的に更新する機能が動作しません。
   
3. ファイル内の既定の **body** クラス (`body {}`) を削除し、次の CSS コードをコピーしてファイルに貼り付けます。
    
    ```css
    /* The following classes describe the common theme information for office documents */ 

    /* Basic Font and Background Colors for text */ 
    .office-docTheme-primary-fontColor { color:#000000; } 
    .office-docTheme-primary-bgColor { background-color:#ffffff; } 
    .office-docTheme-secondary-fontColor { color: #000000; } 
    .office-docTheme-secondary-bgColor { background-color: #ffffff; } 

    /* Accent color definitions for fonts */ 
    .office-contentAccent1-color { color:#5b9bd5; } 
    .office-contentAccent2-color { color:#ed7d31; } 
    .office-contentAccent3-color { color:#a5a5a5; } 
    .office-contentAccent4-color { color:#ffc000; } 
    .office-contentAccent5-color { color:#4472c4; } 
    .office-contentAccent6-color { color:#70ad47; } 

    /* Accent color for backgrounds */ 
    .office-contentAccent1-bgColor { background-color:#5b9bd5; } 
    .office-contentAccent2-bgColor { background-color:#ed7d31; } 
    .office-contentAccent3-bgColor { background-color:#a5a5a5; } 
    .office-contentAccent4-bgColor { background-color:#ffc000; } 
    .office-contentAccent5-bgColor { background-color:#4472c4; } 
    .office-contentAccent6-bgColor { background-color:#70ad47; } 

    /* Accent color for borders */ 
    .office-contentAccent1-borderColor { border-color:#5b9bd5; } 
    .office-contentAccent2-borderColor { border-color:#ed7d31; } 
    .office-contentAccent3-borderColor { border-color:#a5a5a5; } 
    .office-contentAccent4-borderColor { border-color:#ffc000; } 
    .office-contentAccent5-borderColor { border-color:#4472c4; } 
    .office-contentAccent6-borderColor { border-color:#70ad47; } 

    /* links */ 
    .office-a { color: #0563c1; } 
    .office-a:visited { color: #954f72; } 

    /* Body Fonts */ 
    .office-bodyFont-eastAsian { } /* East Asian name of the Font */ 
    .office-bodyFont-latin { font-family:"Calibri"; } /* Latin name of the Font */ 
    .office-bodyFont-script { } /* Script name of the Font */ 
    .office-bodyFont-localized { font-family:"Calibri"; } /* Localized name of the Font. Corresponds to the default font of the culture currently used in Office.*/ 

    /* Headers Font */ 
    .office-headerFont-eastAsian { } 
    .office-headerFont-latin { font-family:"Calibri Light"; } 
    .office-headerFont-script { } 
    .office-headerFont-localized { font-family:"Calibri Light"; } 

    /* The following classes define font and background colors for Office UI themes. These classes should only be used in task pane add-ins */ 

    /* Basic Font and Background Colors for PPT */ 
    .office-officeTheme-primary-fontColor { color:#b83b1d; } 
    .office-officeTheme-primary-bgColor { background-color:#dedede; } 
    .office-officeTheme-secondary-fontColor { color:#262626; } 
    .office-officeTheme-secondary-bgColor { background-color:#ffffff; }
    ```
4. Visual Studio 以外のツールを使用して アドイン を作成している場合は、手順 3 の CSS コードをテキスト ファイルにコピーし、OfficeThemes.css としてファイルを保存したことを確認してください。   

## <a name="reference-officethemescss-in-your-add-ins-html-pages"></a>アドインの HTML ページで OfficeThemes.css を参照する

アドイン プロジェクトで OfficeThemes.css ファイルを使用するには、この形式でアドインの UI を実装する Web ページ `<link>` (.html、.aspx、.php ファイルなど) のタグ内に OfficeThemes.css ファイルを参照するタグを追加します。 `<head>`

```HTML
<link href="<local_path_to_OfficeThemes.css>" rel="stylesheet" type="text/css" />
```

Visual Studio でこの作業を行うには、次の手順に従ってください。

1. [**新規プロジェクトの作成**] を選択します。

2. 検索ボックスを使用して、**アドイン** と入力します。 [**PowerPoint Web アドイン**] を選択し、[**次へ**] を選択します。

3. プロジェクトに名前を付けて、[**作成**] を選択します。

3. **[Office アドインの作成]** ダイアログ ウィンドウで、**[新機能を PowerPoint に追加する]** を選択してから、**[完了]** を選択してプロジェクトを作成します。

4. Visual Studio によってソリューションとその 2 つのプロジェクトが作成され、**ソリューション エクスプローラー** に表示されます。**Home.html** ファイルが Visual Studio で開かれます。

5. 既定のテンプレートの Home.html など、アドインの UI を実装する HTML ページで `<link>` 、OfficeThemes.css ファイルを参照するタグ内に次のタグを追加します。 `<head>`
    
    ```HTML
    <link href="../../Content/OfficeThemes.css" rel="stylesheet" type="text/css" />
    ```

Visual Studio 以外のツールでアドインを作成している場合は、アドインに展開する OfficeThemes.css のコピーへの相対パスを指定して、同じ形式の `<link>` タグを追加してください。

### <a name="use-officethemescss-document-theme-classes-in-your-content-add-ins-html-page"></a>コンテンツ アドインの HTML ページで OfficeThemes.css ドキュメント テーマ クラスを使用する

OfficeTheme.css 文書テーマ クラスを使用するコンテンツ アドイン 内の HTML の簡単な例を以下に示します。文書のテーマで使用される 12 色と 2 つのフォントに対応する OfficeThemes.css クラスの詳細については、「 [コンテンツ アドインのテーマ クラス](#theme-classes-for-content-add-ins)」を参照してください。

```HTML
<body>
    <div id="themeSample" class="office-docTheme-primary-fontColor ">
        <h1 class="office-headerFont-latin">Hello world!</h1> 
        <h1 class="office-headerFont-latin office-contentAccent1-bgColor">Hello world!</h1> 
        <h1 class="office-headerFont-latin office-contentAccent2-bgColor">Hello world!</h1> 
        <h1 class="office-headerFont-latin office-contentAccent3-bgColor">Hello world!</h1> 
        <h1 class="office-headerFont-latin office-contentAccent4-bgColor">Hello world!</h1> 
        <h1 class="office-headerFont-latin office-contentAccent5-bgColor">Hello world!</h1> 
        <h1 class="office-headerFont-latin office-contentAccent6-bgColor">Hello world!</h1> 
        <p class="office-bodyFont-latin office-docTheme-secondary-fontColor">Hello world!</p> 
    </div>
</body>
```

実行時に、既定のテーマテーマを使用するプレゼンテーションに挿入 **Office** コンテンツ アドインは次のようにレンダリングされます。

![テーマを使用して実行Officeアプリ。](../images/office15-app-content-app-office-theme.png)

別のテーマを使用するようにプレゼンテーションを変更するか、プレゼンテーションのテーマをカスタマイズすると、OfficeThemes.css クラスで指定されたフォントと色は、プレゼンテーションのテーマのフォントと色に対応するように動的に更新されます。前述の例と同じ HTML を使用すると、アドインの挿入先のプレゼンテーションで **ファセット** のテーマが使用され、アドインは次のように表示されます。

![ファセット テーマで実行されているコンテンツ アプリ。](../images/office15-app-content-app-facet-theme.png)

### <a name="use-officethemescss-office-ui-theme-classes-in-your-task-pane-add-ins-html-page"></a>作業ウィンドウ アドインの HTML ページOffice UI テーマ クラスに OfficeThemes.css を使用する

ユーザーは、文書のテーマに加えて、すべての Office アプリケーションの Office ユーザー インターフェイスの配色をカスタマイズできます。そのためには、[**ファイル**]  >  [**アカウント**]  >  [**Office テーマ**] ドロップ ダウン ボックスを使用します。

OfficeTheme.css クラスを使用してフォントの色と背景色を指定する作業ウィンドウ アドイン 内の HTML の簡単な例を次に示します。Office UI テーマのフォントと色に対応する OfficeThemes.css クラスの詳細については、「[作業ウィンドウ アドインのテーマ クラス](#theme-classes-for-task-pane-add-ins)」を参照してください。

```HTML
<body> 
    <div id="content-header" class="office-officeTheme-primary-fontColor office-officeTheme-primary-bgColor"> 
        <div class="padding">
            <h1>Welcome</h1>
        </div> 
    </div> 
    <div id="content-main" class="office-officeTheme-secondary-fontColor office-officeTheme-secondary-bgColor"> 
        <div class="padding"> 
            <p>Add home screen content here.</p> 
            <p>For example:</p> 
            <button id="get-data-from-selection">Get data from selection</button> 
            <p><a target="_blank" class="office-a" href="https://go.microsoft.com/fwlink/?LinkId=276812">Find more samples online...</a></p>
        </div>
    </div>
</body> 
```

PowerPoint で [**ファイル**]  >  [**アカウント**]  >  [**Office テーマ**] を [**白**] に設定して実行すると、作業ウィンドウ アドインは次のように表示されます。

![[ホワイト] テーマOffice作業ウィンドウ。](../images/office15-app-task-pane-theme-white.png)

**[Office テーマ]** を **[濃い灰色]** に変更すると、OfficeThemes.css クラスで指定されたフォントと色は動的に更新されて次のように表示されます。

![濃い灰色のテーマOffice作業ウィンドウ。](../images/office15-app-task-pane-theme-dark-gray.png)

## <a name="officethemecss-classes"></a>OfficeTheme.css のクラス

OfficeThemes.css には、PowerPoint 用のコンテンツ アドインおよび作業ウィンドウ アドインと併用できる 2 つのクラスのセットが含まれています。

### <a name="theme-classes-for-content-add-ins"></a>コンテンツ アドインのテーマ クラス

OfficeThemes.css ファイルには、文書テーマで使用される 2 つのフォントと 12 色に対応するクラスがあります。これらのクラスの適切な使用法は、PowerPoint 用のコンテンツ アドインと併用して、アドインのフォントと色が挿入先のプレゼンテーションに合わせて調整されるようにすることです。

### <a name="theme-fonts-for-content-add-ins"></a>コンテンツのアドインのテーマのフォント

|**クラス**|**説明**|
|:-----|:-----|
| `office-bodyFont-eastAsian`|本文のフォントの東アジア言語の名前。|
| `office-bodyFont-latin`|本文のフォントのラテン文字の名前。既定は「Calabri」です。|
| `office-bodyFont-script`|本文のフォントのスクリプト名。|
| `office-bodyFont-localized`|本文のフォントのローカライズされた名前。Office で現在使用されているカルチャに従って既定のフォント名を指定します。|
| `office-headerFont-eastAsian`|ヘッダーのフォントの東アジア言語の名前。|
| `office-headerFont-latin`|ヘッダーのフォントのラテン文字の名前。既定は「Calabri Light」です。|
| `office-headerFont-script`|ヘッダーのフォントのスクリプト名。|
| `office-headerFont-localized`|ヘッダーのフォントのローカライズされた名前。Office で現在使用されているカルチャに従って既定のフォント名を指定します。|

### <a name="theme-colors-for-content-add-ins"></a>コンテンツのアドインのテーマの色

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

![既定のOfficeテーマの色の例です。](../images/office15-app-default-office-theme-colors.png)

### <a name="theme-classes-for-task-pane-add-ins"></a>作業ウィンドウ アドインのテーマ クラス

OfficeThemes.css ファイルには、Office アプリケーション UI テーマで使用されるフォントと背景に割り当てられた 4 色に対応するクラスがあります。これらのクラスの適切な使用法は、PowerPoint 用の作業ウィンドウ アドインと併用して、アドインの色が Office 内の他の組み込み作業ウィンドウに合わせて調整されるようにすることです。

### <a name="theme-font-and-background-colors-for-task-pane-add-ins"></a>作業ウィンドウ アドインのテーマのフォントと背景色

|**クラス**|**説明**|
|:-----|:-----|
| `office-officeTheme-primary-fontColor`|第 1 フォントの色。既定は # B83B1D です。|
| `office-officeTheme-primary-bgColor`|第 1 背景色。既定は #DEDEDE です。|
| `office-officeTheme-secondary-fontColor`|第 2 フォントの色。既定は 262626 です。|
| `office-officeTheme-secondary-bgColor`|第 2 背景色。既定は #FFFFFF です。|

## <a name="see-also"></a>関連項目

- [PowerPoint 用のコンテンツ アドインと作業ウィンドウ アドインを作成する](../powerpoint/powerpoint-add-ins.md)

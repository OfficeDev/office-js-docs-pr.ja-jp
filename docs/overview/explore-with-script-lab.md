---
title: Script Lab を使用して Office JavaScript API を探索する
description: Script Lab を使用して、Office JS API およびプロトタイプの機能を調べます。
ms.date: 04/16/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 6fb886f1c86267ed7081d1892d1314798ab4cedc
ms.sourcegitcommit: 803587b324fc8038721709d7db5664025cf03c6b
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/17/2020
ms.locfileid: "43547256"
---
# <a name="explore-office-javascript-api-using-script-lab"></a>Script Lab を使用して Office JavaScript API を探索する

AppSource から無料で入手できる [Script Lab アドイン](https://appsource.microsoft.com/product/office/WA104380862)を使用すると、Excel や Word などの Office プログラムでの作業中に Office JavaScript API を調査できます。 Script Lab は、アドインで必要な機能のプロトタイプを作成して検証するときに、開発ツールキットに追加する便利なツールです。

## <a name="what-is-script-lab"></a>Script Lab とは

Script Lab は、Excel、Word、または PowerPoint で Office JavaScript API を使用して Office アドインを開発する方法を学習したい人のためのツールです。 IntelliSense を提供しているので、何が利用できるのかを見ることができ、Visual Studio Code で使用されているのと同じフレームワークである Monaco フレームワークの上に構築されています。 Script Lab では、サンプルのライブラリにアクセスして、簡単に機能を試すことができます。また、独自のコードの開始点としてサンプルを使用することもできます。 Script Lab を使用して、プレビュー API を試すこともできます。

今のところいいですか? この 1 分間のビデオを見て、Script Lab の動作を確認します。

[![Excel、Word、PowerPoint での Script Lab の実行を紹介するプレビュー ビデオ。](../images/screenshot-wide-youtube.png 'Script Lab のプレビュー ビデオ')](https://aka.ms/scriptlabvideo)

## <a name="key-features"></a>主な機能

Script Lab には、Office JavaScript API およびプロトタイプ アドインの機能の調査に役立つ機能が多数用意されています。

### <a name="explore-samples"></a>サンプルの確認

API を使用してタスクを完了する方法を示す組み込みのサンプル スニペットのコレクションを使用してすぐに開始できます。 サンプルを実行すると、作業ウィンドウまたはドキュメントですばやく結果を表示したり、API のしくみをサンプルで確認して学んだり、独自のアドインのプロトタイプにサンプルを使用したりもできます。

![サンプル](../images/script-lab-samples.jpg)

### <a name="code-and-style"></a>コードとスタイル

Office JS API を呼び出す JavaScript または TypeScript コードに加えて、各スニペットには、作業ウィンドウのコンテンツを定義する HTML マークアップと、作業ウィンドウの外観を定義する CSS も含まれています。 HTML マークアップと CSS をカスタマイズして、独自のアドインの作業ウィンドウ デザインのプロトタイプを作成する際に、要素の配置とスタイル設定を試すことができます。

> [!TIP]
> スニペット内でプレビュー API を呼び出すには、スニペットのライブラリを更新して、ベータ CDN (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) とプレビューの種類の定義 `@types/office-js-preview` を使用する必要があります。 また、一部のプレビュー API は、[Office Insider プログラム](https://insider.office.com)にサインアップして、Insider ビルドの Office を実行している場合にのみアクセスできます。

### <a name="save-and-share-snippets"></a>スニペットの保存と共有

既定では、Script Lab で開いたスニペットはブラウザーのキャッシュに保存されます。 スニペットを完全に保存するには、そのスニペットを [GitHub の Gist](https://gist.github.com) にエクスポートします。 自分専用にスニペットを保存するには、秘密の Gist を作成するか、他のユーザーと共有する予定がある場合はパブリックの Gist を作成します。

![共有オプション](../images/script-lab-share.jpg)

### <a name="import-snippets"></a>スニペットのインポート

スニペット YAML が保存されているパブリック [ GitHub の Gist ](https://gist.github.com) に URL を指定するか、スニペットの完全な YAML を貼り付けて、スニペットを Script Lab にインポートできます。 この機能は、GitHub の Gist にスニペットを公開するか、スニペットの YAML を提供すると、他のユーザーがスニペットを自分と共有しているシナリオで役立ちます。

![スニペットのインポート オプション](../images/script-lab-import-snippet.jpg)

## <a name="supported-clients"></a>サポートされるクライアント

Script Lab は、次のクライアント上の Excel、Word、PowerPoint でサポートされています。

- Windows での Office 2013 以降
- Mac での Office 2016 以降
- Office on the web

## <a name="next-steps"></a>次の手順

Excel、Word、または PowerPoint で Script Lab を使用するには、AppSource から [Script Lab アドイン](https://appsource.microsoft.com/product/office/WA104380862)をインストールします。 

新しいスニペットを [office-js-snippets](https://github.com/OfficeDev/office-js-snippets#office-js-snippets) GitHub リポジトリに投稿し、Script Lab のサンプル ライブラリを拡張してください。

最初の Office アドインを作成する準備ができたら、[Excel](../quickstarts/excel-quickstart-jquery.md)、[Outlook](../quickstarts/outlook-quickstart.md)、[Word](../quickstarts/word-quickstart.md)、[OneNote ](../quickstarts/onenote-quickstart.md)、[PowerPoint](../quickstarts/powerpoint-quickstart.md)、または [Project](../quickstarts/project-quickstart.md) のクイック スタートを試してください。

## <a name="see-also"></a>関連項目

- [Script Lab を取得する](https://appsource.microsoft.com/product/office/WA104380862)
- [Script Lab の詳細情報](https://github.com/OfficeDev/script-lab#script-lab-a-microsoft-garage-project)
- [Office 365 Developer Program に参加する](https://developer.microsoft.com/office/dev-program)
- [Office アドインを構築する](../overview/office-add-ins-fundamentals.md)

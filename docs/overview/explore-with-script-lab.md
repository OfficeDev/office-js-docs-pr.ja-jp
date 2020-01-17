---
title: Script Lab を使用して Office JavaScript API を探索する
description: スクリプトラボを使用して、Office JS API とプロトタイプ機能を調査します。
ms.date: 07/05/2019
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Normal
ms.openlocfilehash: 3212aec08cdf4e0185ae5856ae522b1d81e28ea1
ms.sourcegitcommit: 212c810f3480a750df779777c570159a7f76054a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/17/2020
ms.locfileid: "41216974"
---
# <a name="explore-office-javascript-api-using-script-lab"></a>Script Lab を使用して Office JavaScript API を探索する

[Script Lab アドイン](https://appsource.microsoft.com/product/office/WA104380862)は appsource から無料で利用できます。これにより、Excel や Word などの office プログラムで作業しているときに OFFICE JavaScript API を調べることができます。 スクリプトラボは、アドインに必要な機能を試作して検証する際に開発ツールキットに追加する便利なツールです。

## <a name="what-is-script-lab"></a>スクリプトラボとは

スクリプトラボは、Excel、Word、または PowerPoint で Office JavaScript API を使用して Office アドインを開発する方法について学習する必要があるユーザーのためのツールです。 これにより IntelliSense が提供され、Visual Studio Code で使用されるのと同じフレームワークである、使用可能なものと、モナコフレームワークに基づいて構築されているものがわかります。 スクリプトラボを使用すると、サンプルのライブラリにアクセスして、機能をすばやく試すことができます。また、サンプルを独自のコードの開始点として使用することもできます。 スクリプトラボを使用してプレビュー Api を試すこともできます。

これまでに良好なことがありますか? この1分間のビデオを見て、実行中のスクリプトラボを確認してください。

[![Excel、Word、PowerPoint での Script Lab の実行を紹介するプレビュー ビデオ。](../images/screenshot-wide-youtube.png 'Script Lab のプレビュー ビデオ')](https://aka.ms/scriptlabvideo)

## <a name="key-features"></a>主な機能

スクリプトラボ Office JavaScript API と prototype アドインの機能について調べるのに役立つさまざまな機能が用意されています。

### <a name="explore-samples"></a>サンプルを検索する

API を使用してタスクを実行する方法を示す組み込みのサンプルスニペットのコレクションを使用して、すぐに作業を開始できます。 サンプルを実行すると、作業ウィンドウまたはドキュメントの結果をすぐに確認したり、サンプルを調べて API のしくみを確認したり、サンプルを使用して独自のアドインをプロトタイプしたりすることもできます。

![サンプル](../images/script-lab-samples.jpg)

### <a name="code-and-style"></a>コードとスタイル

Office JS API を呼び出す JavaScript または TypeScript コードに加えて、各スニペットには、作業ウィンドウの外観を定義する、作業ウィンドウと CSS のコンテンツを定義する HTML マークアップも含まれています。 HTML マークアップと CSS をカスタマイズして、独自のアドインの作業ウィンドウデザインを試作する際に、要素の配置とスタイル設定を試すことができます。

> [!TIP]
> スニペット内でプレビュー Api を呼び出すには、スニペットのライブラリを更新して、ベータ CDN (`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`) とプレビューの種類の定義`@types/office-js-preview`を使用する必要があります。 また、一部のプレビュー Api は、 [Office insider プログラム](https://products.office.com/office-insider)にサインアップし、Office の insider ビルドを実行している場合にのみアクセスできます。

### <a name="save-and-share-snippets"></a>スニペットの保存と共有

既定では、スクリプトラボで開いたスニペットはブラウザーのキャッシュに保存されます。 スニペットを完全に保存するには、それを[GitHub gist](https://gist.github.com)にエクスポートします。 独自にスニペットを保存するための secret gist を作成したり、他のユーザーと共有する予定がある場合は、パブリックな gist を作成したりします。

![共有オプション](../images/script-lab-share.jpg)

### <a name="import-snippets"></a>スニペットのインポート

スニペットをスクリプトラボにインポートするには、スニペット YAML が格納されているパブリック[GitHub gist](https://gist.github.com)への URL を指定するか、スニペットの完全な yaml に貼り付けます。 この機能は、他のユーザーが自分のスニペットを GitHub gist に発行するか、スニペットの YAML を提供することによって、自分のスニペットを共有しているシナリオで役立つことがあります。

![スニペットのインポートオプション](../images/script-lab-import-snippet.jpg)

## <a name="supported-clients"></a>サポートされるクライアント

スクリプトラボは、Excel、Word、および PowerPoint の次のクライアントでサポートされています。

- Office 2013 以降 (Windows)
- Office 2016 以降の Mac
- Office on the web

## <a name="next-steps"></a>次のステップ

Excel、Word、または PowerPoint でスクリプトラボを使用するには、AppSource から[スクリプトラボアドイン](https://appsource.microsoft.com/product/office/WA104380862)をインストールします。 

[Office js](https://github.com/OfficeDev/office-js-snippets#office-js-snippets)の GitHub リポジトリに新しいスニペットを投稿することによって、スクリプトラボのサンプルライブラリを拡張することをお歓迎します。

最初の Office アドインを作成する準備ができたら、 [Excel](../quickstarts/excel-quickstart-jquery.md)、 [Outlook](/outlook/add-ins/quick-start?context=office/dev/add-ins/context)、 [Word](../quickstarts/word-quickstart.md)、 [OneNote](../quickstarts/onenote-quickstart.md)、 [PowerPoint](../quickstarts/powerpoint-quickstart.md)、または[Project](../quickstarts/project-quickstart.md)のクイックスタートをお試しください。

## <a name="see-also"></a>関連項目

- [スクリプトラボの取得](https://appsource.microsoft.com/product/office/WA104380862)
- [スクリプトラボの詳細情報](https://github.com/OfficeDev/script-lab#script-lab-a-microsoft-garage-project)
- [Office 365 Developer Program に参加する](https://developer.microsoft.com/office/dev-program)
- [Office アドインを構築する](../overview/office-add-ins-fundamentals.md)

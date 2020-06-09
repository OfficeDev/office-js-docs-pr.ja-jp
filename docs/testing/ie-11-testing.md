---
ms.date: 05/16/2020
description: Internet Explorer 11 を使用して Office アドインをテストします。
title: Internet Explorer 11 のテスト
localization_priority: Normal
ms.openlocfilehash: 4ea2b4da153e2908f928086cd4997502c194e578
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611205"
---
# <a name="test-your-office-add-in-using-internet-explorer-11"></a>Internet Explorer 11 を使用して Office アドインをテストする

アドインの仕様によっては、以前のバージョンの Windows および Office をサポートすることを計画している場合があります。これには、Internet Explorer 11 でのテストが必要になります。 これは、多くの場合、アドインを AppSource に提出する際に必要になります。 このテストでは、次のコマンドラインツールを使用して、アドインで使用されるより新しいランタイムを Internet Explorer 11 ランタイムに切り替えることができます。

## <a name="pre-requisites"></a>前提条件

- [Node.js](https://nodejs.org/) (最新 [LTS](https://nodejs.org/about/releases) バージョン)
- コード エディター。 [Visual Studio コード](https://code.visualstudio.com/)をお勧めします。
- [Office Insider program の一部である](https://insider.office.com)

これらの手順では、その前に Yo Office ジェネレータープロジェクトを設定していることを前提としています。 これを実行していない場合は、「 [Excel アドインの](../quickstarts/excel-quickstart-jquery.md)場合」などのクイックスタートを読むことを検討してください。

## <a name="using-ie11-tooling"></a>IE11 ツールを使用する

1. Yo Office ジェネレータープロジェクトを作成します。 選択するプロジェクトの種類に関係なく、このツールはすべてのプロジェクトの種類で機能します。

> !こと既存のプロジェクトがあり、新しいプロジェクトを作成せずにこのツールを追加する場合は、この手順をスキップして次の手順に進みます。 

2. 新しいプロジェクトのルートフォルダーで、コマンドラインで次のコマンドを実行します。

```command&nbsp;line
office-add-dev-settings webview manifest.xml ie
```
Web ビューの種類が IE に設定されていることを示すメモがコマンドラインに表示されます。

> !部このツールを使用する必要はありませんが、Internet Explorer 11 ランタイムに関連する問題の大部分をデバッグするのに役立ちます。 堅牢性を完全にするには、Windows 7 および Office 2013 のコピーがインストールされたコンピューターを使用してテストする必要があります。

## <a name="command-settings"></a>コマンドの設定

マニフェストパスが異なる場合は、次のようにコマンドでこれを指定します。

`office-add-dev-settings webview [path to your manifest] ie`

また、このコマンドは、 `office-addin-dev-settings webview` 引数としていくつかのランタイムを取ることができます。

- internet
- 下辺
- 既定値です。

## <a name="see-also"></a>関連項目
* [Office アドインのテストとデバッグ](test-debug-office-add-ins.md)
* [テスト用に Office アドインをサイドロードする](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [Windows 10 で開発者ツールを使用してアドインをデバッグする](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
* [作業ウィンドウからデバッガーをアタッチする](attach-debugger-from-task-pane.md)
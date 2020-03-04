---
ms.date: 02/20/2020
title: ブラウザーのランタイムを共有するように Excel アドインを構成する (プレビュー)
ms.prod: excel
description: Excel アドインを構成して、ブラウザーのランタイムを共有し、同じランタイムでリボン、作業ウィンドウ、カスタム関数のコードを実行できるようにします。
localization_priority: Priority
ms.openlocfilehash: 7945bd8fdb29a9d6d44d7d29676410a54bacf83f
ms.sourcegitcommit: dd6d00202f6466c27418247dad7bd136555a6036
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/26/2020
ms.locfileid: "42284133"
---
# <a name="configure-your-excel-add-in-to-use-a-shared-javascript-runtime-preview"></a>共有の JavaScript ランタイムを使用するように Excel アドインを構成する (プレビュー)

[!include[Running custom functions in a shared runtime note](../includes/excel-shared-runtime-preview-note.md)]

Windows または Mac で Excel を実行する場合、アドインは、リボン ボタン、カスタム関数、作業ウィンドウのコードを別の JavaScript ランタイム環境で実行します。 これにより、グローバル データを簡単に共有できない、カスタム関数からすべての CORS 機能にアクセスできないなどの制限が発生します。

ただし、Excel アドインを構成すれば、共有の JavaScript ランタイムでコードを共有できるようになります。 これにより、アドイン間での調整が容易になり、アドインのすべての部分から DOM や CORS にアクセスできます。 また、ドキュメントを開いているときにコードを実行したり、作業ウィンドウが閉じた状態でコードを実行したりできます。 共有ランタイムが使用できるようにアドインを構成するには、この記事の手順に従います。

## <a name="create-the-add-in-project"></a>アドイン プロジェクトの作成

新しいプロジェクトを開始する場合は、次の手順に従って、Yeoman ジェネレーターを使って Excel アドインを作成します。 次のコマンドを実行し、プロンプトに次の回答を入力します。

```command line
yo office
```

- プロジェクトの種類を選択する:  **Excel カスタム関数アドイン プロジェクト**
- スクリプトの種類を選択する:  **JavaScript**
- アドインの名前を何にしますか?  **個人用 Office アドイン**

![アドイン プロジェクトを作成するための Office からのプロンプトへ応答するスクリーンショット。](../images/yo-office-excel-project.png)

ウィザードを完了すると、ジェネレーターによってプロジェクトが作成され、サポートしているノード コンポーネントがインストールされます。

## <a name="configure-the-manifest"></a>マニフェストを構成する

新規または既存のプロジェクトで共有ランタイムが使用できるように構成するには、次の手順を実行します。

1. Visual Studio Code を開始して [**個人用 Office アドイン**] プロジェクトを開きます。
2. 
            **manifest.xml** ファイルを開きます。
3. `<VersionOverrides>` セクションを探し、次の `<Runtimes>` セクションを追加します。 作業ウィンドウを閉じてもカスタム関数が引き続き機能するように、有効期間は**長く**する必要があります。 resid は `ContosoAddin.Url` で、後述のリソースのセクションの文字列を参照します。 resid には任意の値を使用できますが、アドイン要素のその他の要素の resid と一致している必要があります。

   ```xml
   <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
     <Hosts>
       <Host xsi:type="Workbook">
       <Runtimes>
         <Runtime resid="ContosoAddin.Url" lifetime="long" />
       </Runtimes>
       <AllFormFactors>
   ```

4. `<Page>` 要素で、ソースの場所を **Functions.Page.Url** から **ContosoAddin.Url** に変更します。 この resid は、`<Runtime>` resid の要素と一致しています。 カスタム関数がない場合は、**Page** エントリがないため、この手順は省略できます。

   ```xml
   <AllFormFactors>
   ...
   <Page>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Page>
   ...
   ```

5. `<DesktopFormFactor>` セクションで、**FunctionFile** を **Commands.Url** から **ContosoAddin.Url** を使用するように変更します。 アクション コマンドがない場合は、**FunctionFile** エントリがないため、この手順は省略できます。

   ```xml
   <DesktopFormFactor>
   <GetStarted>
   ...
   </GetStarted>
   <FunctionFile resid="ContosoAddin.Url"/>
   ```

6. `<Action>` セクションで、ソースの場所を **Taskpane.Url** から **ContosoAddin.Url** に変更します。 作業ウィンドウがない場合は、**ShowTaskpane** アクションがないため、この手順は省略できます。

   ```xml
   <Action xsi:type="ShowTaskpane">
   <TaskpaneId>ButtonId1</TaskpaneId>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Action>
   ```

7. **taskpane.html** を指す **ContosoAddin.Url** の新しい **Url ID** を追加します。

   ```xml
   <bt:Urls>
   <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
   ...
   <bt:Url id="ContosoAddin.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
   ...
   ```

8. 変更を保存してプロジェクトを再ビルドします。

   ```command line
   npm run build
   ```

## <a name="runtime-lifetime"></a>ランタイムの有効期間

`Runtime` 要素を追加するときに、有効期間も `long` または `short` の値で指定します。 この値を `long` に設定すると、ドキュメントを開くとアドインを起動したり、作業ウィンドウを閉じた後にコードを継続して実行したり、カスタム関数から CORS および DOM を使用したりできます。

この値を `short`に設定すると、アドインは既定の動作と同様に動作します。 いずれかのリボン ボタンを押すとアドインが開始しますが、リボン ハンドラーの実行を完了した後で、アドインが終了する場合があります。 同様に、作業ウィンドウを開くとアドインが起動します。ただし、作業ウィンドウを閉じると、アドインが終了する場合があります。

```xml
<Runtimes>
  <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="multiple-task-panes"></a>複数の作業ウィンドウ

共有ランタイムを使用する予定がある場合は、複数の作業ウィンドウを使用するようにアドインを設計しないでください。 共有ランタイムは、1 つの作業ウィンドウのみサポートします。 `<TaskpaneID>` のない作業ウィンドウは、別の作業ウィンドウとして扱われますのでご注意ください。

## <a name="next-steps"></a>次の手順

次の記事を参照して、共有ランタイムの機能をいくつかお試しください。

- [カスタム関数から Excel API を呼び出す](call-excel-apis-from-custom-function.md)

## <a name="see-also"></a>関連項目

- [概要: 共有の JavaScript ランタイムでアドイン コードを実行する (プレビュー)](custom-functions-shared-overview.md)

---
ms.date: 05/17/2020
title: ブラウザーランタイムを共有するように Excel アドインを構成する
ms.prod: excel
description: Excel アドインを構成して、ブラウザーのランタイムを共有し、同じランタイムでリボン、作業ウィンドウ、カスタム関数のコードを実行できるようにします。
localization_priority: Priority
ms.openlocfilehash: 166e0f9f2c2fe99b303484d5212e5d8b23726d6a
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/18/2020
ms.locfileid: "44276016"
---
# <a name="configure-your-excel-add-in-to-use-a-shared-javascript-runtime"></a>共有された JavaScript ランタイムを使用するように Excel アドインを構成する

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Windows または Mac で Excel を実行する場合、アドインは、リボン ボタン、カスタム関数、作業ウィンドウのコードを別の JavaScript ランタイム環境で実行します。 これにより、グローバルデータを簡単に共有することができず、カスタム関数からすべての CORS 機能にアクセスできないなどの制限が生じます。

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

>!こと既定の有効期間値はです `short` が、Excel アドインで使用することをお勧め `long` します。この例でランタイムをに設定した場合 `short` 、Excel アドインは、いずれかのリボンボタンが押されたときに開始しますが、リボンハンドラーの実行が終了すると、シャットダウンする可能性があります。 同様に、作業ウィンドウを開くとアドインが起動します。ただし、作業ウィンドウを閉じると、アドインが終了する場合があります。

```xml
<Runtimes>
  <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="multiple-task-panes"></a>複数の作業ウィンドウ

共有ランタイムを使用することを計画している場合は、複数の作業ウィンドウを使用するようにアドインをデザインしないでください。 共有ランタイムは、1つの作業ウィンドウの使用のみをサポートしています。 `<TaskpaneID>` のない作業ウィンドウは、別の作業ウィンドウとして扱われますのでご注意ください。

## <a name="next-steps"></a>次のステップ

- Excel JavaScript Api の使用および共有ランタイムでの Excel のカスタム関数の使用方法の詳細については、「[カスタム関数から Excel API を呼び出す](call-excel-apis-from-custom-function.md)」の記事を参照してください。
- [パターンアンドプラクティス]のサンプル [リボンと作業ウィンドウの UI を管理し、ドキュメント オープンのコードを実行](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-shared-runtime-scenario)を探索して、共有されている JavaScript ランタイムの大規模な例をご覧ください。

## <a name="see-also"></a>関連項目

- [概要: 共有 JavaScript ランタイムでアドインコードを実行する](custom-functions-shared-overview.md)

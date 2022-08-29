---
title: Angular で Office アドインを開発する
description: Angularを使用して、単一ページ アプリケーションとして Office アドインを作成します。
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: bbac0f94b731b2853e17ed3db785b50ea99ef6e4
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/24/2022
ms.locfileid: "67422958"
---
# <a name="develop-office-add-ins-with-angular"></a>Angular で Office アドインを開発する

この記事では、Angular 2+ を使って、単一ページのアプリケーションとして Office アドインを作成する方法を説明します。

> [!NOTE]
> Angularを使用して Office アドインを作成した経験に基づいて、何か貢献する必要がありますか? [GitHub でこの記事に](https://github.com/OfficeDev/office-js-docs-pr/blob/master/docs/develop/add-ins-with-angular2.md)投稿することも、リポジトリに[問題](https://github.com/OfficeDev/office-js-docs-pr/issues)を送信してフィードバックを提供することもできます。

Angular フレームワークを使用してビルドされる Office アドインのサンプルについては、「[Angular でビルドする Word スタイル チェック アドイン](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)」を参照してください。

## <a name="install-the-typescript-type-definitions"></a>TypeScript 型の定義をインストールする

Node.js ウィンドウを開き、コマンド ラインで次のように入力します。

```command&nbsp;line
npm install --save-dev @types/office-js
```

## <a name="bootstrapping-must-be-inside-officeinitialize"></a>ブートス トラップは必ず Office.initialize 内に

Office、Word、または Excel JavaScript API を呼び出す任意のページで、コードで最初に関数 `Office.initialize`を割り当てる必要があります。 (初期化コードがない場合は、関数本体を空の "`{}`" シンボルにすることができますが、関数を未定義のままに `Office.initialize` しないでください。 詳細については、「 [Office アドインを初期化する](initialize-add-in.md)」を参照してください。Office は、Office JavaScript ライブラリを初期化した直後にこの関数を呼び出します。

Office JavaScript ライブラリが最初に初期化されるように **するには`Office.initialize`、割り当てる関数内でAngularブートストラップ コードを呼び出す必要があります**。 以下は、これを行う方法を示した簡単な例です。 このコードは、プロジェクトの main.ts ファイルの中にある必要があります。

```js
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
import { AppModule } from './app.module';

Office.initialize = function () {
  const platform = platformBrowserDynamic();
  platform.bootstrapModule(AppModule);
};
```

## <a name="use-the-hash-location-strategy-in-the-angular-application"></a>Angular アプリケーションで Hash Location Strategy を使う

Hash Location Strategy を指定しないと、アプリケーションでルート間の移動が機能しない可能性があります。2 つの方法のいずれかでこれを行うことができます。1 つ目の方法は、次の例に示すとおり、アプリ モジュールでプロバイダーをロケーションの戦略に指定できます。これは app.module.ts ファイルに入ります。

```js
import { LocationStrategy, HashLocationStrategy } from '@angular/common';
// Other imports suppressed for brevity

@NgModule({
  providers: [
    { provide: LocationStrategy, useClass: HashLocationStrategy },
    // Other providers suppressed
  ],
  // Other module properties suppressed
})
export class AppModule { }
```

独立したルーティング モジュール内で、ルートを定義する場合は、Hash Location Strategy を指定する別の方法があります。ルーティング モジュールの .ts ファイルで、戦略を指定する `forRoot` 関数に構成オブジェクトを渡します。以下にコードの例を示します。

```js
import { RouterModule, Routes } from '@angular/router';
// Other imports suppressed for brevity

const routes: Routes = // route definitions go here

@NgModule({
  imports: [RouterModule.forRoot(routes, { useHash: true })],
  exports: [RouterModule]
})
export class AppRoutingModule { }
```

## <a name="use-the-office-dialog-api-with-angular"></a>Angularで Office ダイアログ API を使用する

Office のアドインの Dialog API を使えば、アドインでは、メイン ページと情報をやりとりできるセミモードレス ダイアログ ボックスで、ページを開けるようになります。通常、これは作業ウィンドウにあります。

[DisplayDialogAsync](/javascript/api/office/office.ui) メソッドは、ダイアログ ボックスで開くべきページの URL を指定するパラメーターを受け取ります。アドインでは、独立した HTML ページ (基本ページとは異なるページ) でこのパラメーターに渡すか、Angular アプリケーションでルートの URL を渡すことができます。

重要な点として、ルートを渡すと、ダイアログ ボックスによって新しいウィンドウとその実行コンテキストが作成されることに注意してください。 ダイアログ ボックスで、この新しいコンテキストに対して基本ページとそのすべての初期化、およびブートストラップ コードを再度実行し、すべての変数が初期値に設定されます。 この手法により、ダイアログ ボックスで、単一ページのアプリケーションの 2 番目のインスタンスが起動します。 ダイアログ ボックス内の変数を変更するコードは、同じ変数の作業ウィンドウのバージョンは変更しません。 同様に、ダイアログ ボックスには独自のセッション ストレージ ( [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) プロパティ) があり、作業ウィンドウのコードからはアクセスできません。  

## <a name="trigger-the-ui-update"></a>UI の更新をトリガーする

Angular アプリでは UI が更新されない場合があります。これは、コード部分が Angular ゾーンの外から実行されるためです。解決策としては、次の例に示すように、ゾーン内にコードを配置します。

```js
import { NgZone } from '@angular/core';

export class MyComponent {
  constructor(private zone: NgZone) { }

  myFunction() {
    this.zone.run(() => {
      // the codes that need update the UI
    });
  }
}
```

## <a name="use-observable"></a>Observable を使用する

Angular は RxJS (JavaScript の事後対応型の拡張機能) を使用し、RxJS は `Observable` と `Observer` のオブジェクトを導入して非同期処理を実装します。このセクションでは、`Observables` の使い方についての概要を簡単に紹介しています。さらに詳細な情報については、[RxJS](https://rxjs-dev.firebaseapp.com/) の公式ドキュメントを参照してください。

`Observable` は、ある意味で `Promise` オブジェクトに似ています。非同期の呼び出しからすぐに返されますが、すぐには解決されない可能性があります。しかし、`Promise` は、単一の値 (配列オブジェクトのことがあります) なのに対し、`Observable` は、オブジェクトの配列 (メンバーが 1 つだけの可能性あり) です。そのため、コードで `concat`、`map`、`filter` などの[配列メソッド](https://www.w3schools.com/jsref/jsref_obj_array.asp)を `Observable` オブジェクトで呼び出すことができます。

### <a name="push-instead-of-pull"></a>プルの代わりにプッシュする

コードは `Promise` オブジェクトを変数に割り当てることによって "プル" しますが、`Observable` オブジェクトは、値を `Observable` に *登録* するオブジェクトに、"プッシュ" します。サブスクライバーは、`Observer` オブジェクトです。プッシュ アーキテクチャの利点は、時間の経過と共に新しいメンバーを `Observable` 配列に追加できることです。新しいメンバーが追加されると、`Observable` に登録されるすべての `Observer` オブジェクトは通知を受信します。

`Observer` は、関数とともに新規の各オブジェクト ("next" オブジェクトと呼ばれる) を処理するように構成されます。(また、エラーと完了の通知に応答するようにも構成されます。例については、次のセクションを参照してください。)このため、`Observable` オブジェクトは、`Promise` オブジェクトよりも幅広いシナリオで使用できます。たとえば、AJAX 呼び出しから `Observable` を返すことに加えて、`Promise` を返し、`Observable` をテキスト ボックスの "変更" イベント ハンドラーなどのイベント ハンドラーから返すことができます。ユーザーがボックスにテキストを入力するたびに、登録されているすべての `Observer` オブジェクトが、最新のテキストや、アプリケーションの現在の状態を入力として使用することによって、すぐに対応します。

### <a name="wait-until-all-asynchronous-calls-have-completed"></a>すべての非同期呼び出しが完了するまで待ちます

一連の `Promise` オブジェクトの各メンバーが解決されるときのみ確実にコールバックが実行されるようにしたい場合は、`Promise.all()` メソッドを使用します。

```js
myPromise.all([x, y, z]).then(
  // TODO: Callback logic goes here
)
```

`Observable` オブジェクトで同じことを行うには、[Observable.forkJoin()](https://github.com/Reactive-Extensions/RxJS/blob/master/doc/api/core/operators/forkjoin.md) メソッドを使います。  

```js
const source = Observable.forkJoin([x, y, z]);

const subscription = source.subscribe(
  x => {
    // TODO: Callback logic goes here
  },
  err => console.log('Error: ' + err),
  () => console.log('Completed')
);
```

## <a name="compile-the-angular-application-using-the-ahead-of-time-aot-compiler"></a>Ahead-of-Time (AOT) コンパイラを使って Angular アプリケーションをコンパイルする

アプリケーションのパフォーマンスは、ユーザー エクスペリエンスの中でも重要度が高いものの 1 つです。 Angular アプリケーションは、ビルド時に Angular Ahead-of-Time (AOT) コンパイラを使用してアプリをコンパイルすることで最適化することができます。 すべてのソース コード (HTML テンプレートと TypeScript) を効率的な JavaScript コードに変換します。 AOT コンパイラを使用してアプリをコンパイルすると、実行時に追加のコンパイルは実行されません。そのため、HTML テンプレートのレンダリングと非同期要求が高速になります。 さらに、Angular コンパイラを配布可能なアプリケーションに含める必要がないため、アプリケーション全体のサイズが小さくなります。

AOT コンパイラを使用するには、`ng build` または `ng serve` コマンドに `--aot` を追加します。

```command&nbsp;line
ng build --aot
ng serve --aot
```

> [!NOTE]
> Angular Ahead-of-Time (AOT) コンパイラの詳細については、[公式ガイド](https://angular.io/guide/aot-compiler)を参照してください。

## <a name="support-internet-explorer-if-youre-dynamically-loading-officejs"></a>Office.jsを動的に読み込む場合は、Internet Explorer をサポートする

アドインが実行されている Windows バージョンと Office デスクトップ クライアントに基づいて、アドインが Internet Explorer 11 を使用している可能性があります。 (詳細については、「[Office アドインで使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)」を参照してください)。Angularは、いくつかの`window.history`API ですが、これらの API は、Windows デスクトップ クライアントで Office アドインを実行するために使用される IE ランタイムでは機能しません。 これらの API が機能しない場合は、アドインが正常に動作しない可能性があります。たとえば、空白の作業ウィンドウが読み込まれる可能性があります。 これを軽減するために、Office.jsはそれらの API を null 化します。 ただし、Office.jsを動的に読み込む場合は、Office.jsの前に AngularJS が読み込まれる可能性があります。 その場合は、アドインの `window.history` **index.html** ページに次のコードを追加して API を無効にする必要があります。

```js
<script type="text/javascript">window.history.replaceState=null;window.history.pushState=null;</script>
```

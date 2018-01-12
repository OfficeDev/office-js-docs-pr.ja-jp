# <a name="tips-for-creating-office-add-ins-with-angular"></a>Angular で Office アドインを作成するためのヒント

この記事では、Angular 2+ を使って、単一ページのアプリケーションとして Office アドインを作成する方法を説明します。

>**注:**Angular を使用して Office アドインを作成した経験を基に、何か投稿する内容がありますか。[GitHub](https://github.com/OfficeDev/office-js-docs) でこの記事に対して投稿するか、リポジトリで[問題](https://github.com/OfficeDev/office-js-docs/issues)を提出することでフィードバックを提出できます。 

Angular フレームワークを使用してビルドされる Office アドインのサンプルについては、「[Angular でビルドする Word スタイル チェック アドイン](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)」を参照してください。

## <a name="install-the-typescript-type-definitions"></a>TypeScript 型の定義をインストールする
nodejs ウィンドウを開き、コマンド ラインで次のように入力します: `npm install --save-dev @types/office-js`。

## <a name="bootstrapping-must-be-inside-officeinitialize"></a>ブートス トラップは必ず Office.initialize 内に

Office、Word、Excel の JavaScript API を呼び出す任意のページで、コードでまずメソッドを `Office.initialize` プロパティに割り当てる必要があります。(初期化コードがない場合は、メソッド本文は空の "`{}`" 記号でも構いませんが、`Office.initialize` プロパティは未定義のままにはできません。詳細については、「[アドインを初期化する](http://dev.office.com/docs/add-ins/develop/understanding-the-javascript-api-for-office#initializing-your-add-in)」を参照してください。)Office の JavaScript ライブラリを初期化すると、すぐに Office でこのメソッドが呼び出されます。

**Angular のブートストラップ コードは `Office.initialize` に割り当てられたメソッドの中で呼び出すことで**、Office の JavaScript ライブラリが最初に初期化されるようにする必要があります。以下は、これを行う方法を示した簡単な例です。このコードは、プロジェクトの main.ts ファイルの中にある必要があります。

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


## <a name="consider-wrapping-fabric-components-with-angular-components"></a>Fabric コンポーネントと Angular コンポーネントとのラッピングについて検討する

アドインには [Office UI Fabric](http://dev.office.com/fabric#/fabric-js) のスタイルを使用することをお勧めしています。Fabric には、[TypeScript に基づいた](https://github.com/OfficeDev/office-ui-fabric-js)バージョンを含む、いくつかのバージョンに由来するコンポーネントが含まれています。Fabric コンポーネントを、Angular のコンポーネントでラッピングすることによってアドインで使用することを検討してください。これを行う方法を説明した例については、「[Angular でビルドする Word スタイル チェック アドイン](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)」を参照してください。たとえば、[fabric.textfield.wrapper](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/shared/office-fabric-component-wrappers/fabric.textfield.wrapper.component.ts) で定義されている Angular コンポーネントで Fabric ファイルの TextField.ts をインポートすると、その場所に Fabric コンポーネントが定義されます。 


## <a name="using-the-office-dialog-api-with-angular"></a>Angular で Office Dialog API を使用する

Office のアドインの Dialog API を使えば、アドインでは、メイン ページと情報をやりとりできるセミモーダル ダイアログ ボックスで、ページを開けるようになります。通常、これは作業ウィンドウにあります。 

[DisplayDialogAsync](http://dev.office.com/reference/add-ins/shared/officeui.displaydialogasync) メソッドは、ダイアログ ボックスで開くべきページの URL を指定するパラメーターを受け取ります。アドインでは、独立した HTML ページ (基本ページとは異なるページ) でこのパラメーターに渡すか、Angular アプリケーションでルートの URL を渡すことができます。 

重要な点として、ルートを渡すと、ダイアログ ボックスによって新しいウィンドウとその実行コンテキストが作成されることに注意してください。ダイアログ ボックスで、この新しいコンテキストに対して基本ページとそのすべての初期化、およびブートストラップ コードを再度実行し、すべての変数が初期値に設定されます。この手法により、ダイアログ ボックスで、単一ページのアプリケーションの 2 番目のインスタンスが起動します。ダイアログ ボックス内の変数を変更するコードは、同じ変数の作業ウィンドウのバージョンは変更しません。同様に、ダイアログ ボックスには、それ自体にセッション ストレージがあり、作業ウィンドウからコードでそこにアクセスすることはできません。  


## <a name="trigger-the-ui-update"></a>UI の更新をトリガーする

Angular アプリでは UI が更新されない場合があります。これは、コード部分が Angular ゾーンの外から実行されるためです。解決策としては、次の例に示すように、ゾーン内にコードを配置します。

```ts
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

## <a name="using-observable"></a>Observable を使用する

Angular は RxJS (JavaScript の事後対応型の拡張機能) を使用し、RxJS は `Observable` と `Observer` のオブジェクトを導入して非同期処理を実装します。このセクションでは、`Observables` の使い方についての概要を簡単に紹介しています。さらに詳細な情報については、[RxJS](http://reactivex.io/rxjs/) の公式ドキュメントを参照してください。

`Observable` は、ある意味で `Promise` オブジェクトに似ています。非同期の呼び出しからすぐに返されますが、すぐには解決されない可能性があります。しかし、`Promise` は、単一の値 (配列オブジェクトのことがあります) なのに対し、`Observable` は、オブジェクトの配列 (メンバーが 1 つだけの可能性あり) です。そのため、コードで `concat`、`map`、`filter` などの[配列メソッド](http://www.w3schools.com/jsref/jsref_obj_array.asp)を `Observable` オブジェクトで呼び出すことができます。 

### <a name="pushing-instead-of-pulling"></a>プルではなくプッシュ

コードは `Promise` オブジェクトを変数に割り当てることによって "プル" しますが、`Observable` オブジェクトは、値を `Observable` に*登録*するオブジェクトに、"プッシュ" します。サブスクライバーは、`Observer` オブジェクトです。プッシュ アーキテクチャの利点は、時間の経過と共に新しいメンバーを `Observable` 配列に追加できることです。新しいメンバーが追加されると、`Observable` に登録されるすべての `Observer` オブジェクトは通知を受信します。 

`Observer` は、関数とともに新規の各オブジェクト ("next" オブジェクトと呼ばれる) を処理するように構成されます。(また、エラーと完了の通知に応答するようにも構成されます。例については、次のセクションを参照してください。)このため、`Observable` オブジェクトは、`Promise` オブジェクトよりも幅広いシナリオで使用できます。たとえば、AJAX 呼び出しから `Observable` を返すことに加えて、`Promise` を返し、`Observable` をテキスト ボックスの "変更" イベント ハンドラーなどのイベント ハンドラーから返すことができます。ユーザーがボックスにテキストを入力するたびに、登録されているすべての `Observer` オブジェクトが、最新のテキストや、アプリケーションの現在の状態を入力として使用することによって、すぐに対応します。 


### <a name="waiting-until-all-asynchronous-calls-have-completed"></a>すべての非同期呼び出しが完了するまで待機する

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


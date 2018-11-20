---
title: Angular で Office アドインを開発する
description: ''
ms.date: 11/02/2018
ms.openlocfilehash: 312317e594024125e2dc86d23840750e48d81e40
ms.sourcegitcommit: c6723a31b48945ca4c466ba016a3dfc7b6267f5c
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/03/2018
ms.locfileid: "25942252"
---
# <a name="develop-office-add-ins-with-angular"></a><span data-ttu-id="c936e-102">Angular で Office アドインを開発する</span><span class="sxs-lookup"><span data-stu-id="c936e-102">Develop Office Add-ins with Angular</span></span>

<span data-ttu-id="c936e-103">この記事では、Angular 2+ を使って、単一ページのアプリケーションとして Office アドインを作成する方法を説明します。</span><span class="sxs-lookup"><span data-stu-id="c936e-103">This article provides guidance for using Angular 2+ to create an Office Add-in as a single page application.</span></span>

> [!NOTE]
> <span data-ttu-id="c936e-p101">Angular を使用して Office アドインを作成した経験を基に、何か投稿する内容がありますか。[GitHub](https://github.com/OfficeDev/office-js-docs) でこの記事に対して投稿するか、リポジトリで[問題](https://github.com/OfficeDev/office-js-docs-pr/issues)を提出することでフィードバックを提出できます。</span><span class="sxs-lookup"><span data-stu-id="c936e-p101">Do you have something to contribute based on your experience using Angular to create Office Add-ins? You can contribute to this article in [GitHub](https://github.com/OfficeDev/office-js-docs) or provide your feedback by submitting an [issue](https://github.com/OfficeDev/office-js-docs-pr/issues) in the repo.</span></span> 

<span data-ttu-id="c936e-106">Angular フレームワークを使用してビルドされる Office アドインのサンプルについては、「[Angular でビルドする Word スタイル チェック アドイン](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c936e-106">For an Office Add-ins sample that's built using the Angular framework, see [Word Style Checking Add-in Built on Angular](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker).</span></span>

## <a name="install-the-typescript-type-definitions"></a><span data-ttu-id="c936e-107">TypeScript 型の定義をインストールする</span><span class="sxs-lookup"><span data-stu-id="c936e-107">Install the TypeScript type definitions</span></span>
<span data-ttu-id="c936e-108">nodejs ウィンドウを開き、コマンド ラインで次のように入力します:</span><span class="sxs-lookup"><span data-stu-id="c936e-108">Open an nodejs window and enter the following at the command line:</span></span> 

```bash
npm install --save-dev @types/office-js
```

## <a name="bootstrapping-must-be-inside-officeinitialize"></a><span data-ttu-id="c936e-109">ブートス トラップは必ず Office.initialize 内に</span><span class="sxs-lookup"><span data-stu-id="c936e-109">Bootstrapping must be inside Office.initialize</span></span>

<span data-ttu-id="c936e-p102">Office、Word、Excel の JavaScript API を呼び出す任意のページで、コードでまずメソッドを `Office.initialize` プロパティに割り当てる必要があります。(初期化コードがない場合は、メソッド本文は空の "`{}`" 記号でも構いませんが、`Office.initialize` プロパティは未定義のままにはできません。詳細については、「[アドインを初期化する](understanding-the-javascript-api-for-office.md#initializing-your-add-in)」を参照してください。)Office の JavaScript ライブラリを初期化すると、すぐに Office でこのメソッドが呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c936e-p102">On any page that calls the Office, Word, or Excel JavaScript APIs, your code must first assign a method to the `Office.initialize` property. (If you have no initialization code, the method body can be just empty "`{}`" symbols, but you must not leave the `Office.initialize` property undefined. For details, see [Initializing your add-in](understanding-the-javascript-api-for-office.md#initializing-your-add-in).) Office calls this method immediately after it has initialized the Office JavaScript libraries.</span></span>

<span data-ttu-id="c936e-p103">**Angular のブートストラップ コードは `Office.initialize` に割り当てられたメソッドの中で呼び出すことで**、Office の JavaScript ライブラリが最初に初期化されるようにする必要があります。以下は、これを行う方法を示した簡単な例です。このコードは、プロジェクトの main.ts ファイルの中にある必要があります。</span><span class="sxs-lookup"><span data-stu-id="c936e-p103">**Your Angular bootstrapping code must be called inside the method that you assign to `Office.initialize`** to ensure that the Office JavaScript libraries have initialized first. The following is a simple example that shows how to do this. This code should be in the main.ts file of the project.</span></span>

```js
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
import { AppModule } from './app.module';

Office.initialize = function () {
  const platform = platformBrowserDynamic();
  platform.bootstrapModule(AppModule);
};
```

## <a name="use-the-hash-location-strategy-in-the-angular-application"></a><span data-ttu-id="c936e-116">Angular アプリケーションで Hash Location Strategy を使う</span><span class="sxs-lookup"><span data-stu-id="c936e-116">Use the hash location strategy in the Angular application</span></span>

<span data-ttu-id="c936e-p104">Hash Location Strategy を指定しないと、アプリケーションでルート間の移動が機能しない可能性があります。2 つの方法のいずれかでこれを行うことができます。1 つ目の方法は、次の例に示すとおり、アプリ モジュールでプロバイダーをロケーションの戦略に指定できます。これは app.module.ts ファイルに入ります。</span><span class="sxs-lookup"><span data-stu-id="c936e-p104">Navigating between routes in the application might not work if you don't specify the hash location strategy. You can do this in one of two ways. First, you can specify a provider for the location strategy in your app module, as shown in the following example. It goes into the app.module.ts file.</span></span>

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

<span data-ttu-id="c936e-p105">独立したルーティング モジュール内で、ルートを定義する場合は、Hash Location Strategy を指定する別の方法があります。ルーティング モジュールの .ts ファイルで、戦略を指定する `forRoot` 関数に構成オブジェクトを渡します。以下にコードの例を示します。</span><span class="sxs-lookup"><span data-stu-id="c936e-p105">If you define your routes in a separate routing module, there is an alternative way to specify the hash location strategy. In your routing module's .ts file, pass a configuration object to the `forRoot` function that specifies the strategy. The following code is an example.</span></span> 

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


## <a name="consider-wrapping-fabric-components-with-angular-components"></a><span data-ttu-id="c936e-124">Fabric コンポーネントと Angular コンポーネントとのラッピングについて検討する</span><span class="sxs-lookup"><span data-stu-id="c936e-124">Consider wrapping Fabric components with Angular components</span></span>

<span data-ttu-id="c936e-p106">アドインには [Office UI Fabric](https://developer.microsoft.com/fabric#/fabric-js) のスタイルを使用することをお勧めしています。Fabric には、[TypeScript に基づいた](https://github.com/OfficeDev/office-ui-fabric-js)バージョンを含む、いくつかのバージョンに由来するコンポーネントが含まれています。Fabric コンポーネントを、Angular のコンポーネントでラッピングすることによってアドインで使用することを検討してください。これを行う方法を説明した例については、「[Angular でビルドする Word スタイル チェック アドイン](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)」を参照してください。たとえば、[fabric.textfield.wrapper](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/shared/office-fabric-component-wrappers/fabric.textfield.wrapper.component.ts) で定義されている Angular コンポーネントで Fabric ファイルの TextField.ts をインポートすると、その場所に Fabric コンポーネントが定義されます。</span><span class="sxs-lookup"><span data-stu-id="c936e-p106">We recommend using [Office UI Fabric](https://developer.microsoft.com/fabric#/fabric-js) styling in your add-in. Fabric includes components that come in several versions, including a version [based on TypeScript](https://github.com/OfficeDev/office-ui-fabric-js). Consider using Fabric components in your add-in by wrapping them in Angular components. For an example that shows you how to do this, see [Word Style Checking Add-in Built on Angular](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker). Note, for example, how the Angular component defined in [fabric.textfield.wrapper](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/shared/office-fabric-component-wrappers/fabric.textfield.wrapper.component.ts) imports the Fabric file TextField.ts, where the Fabric component is defined.</span></span> 


## <a name="using-the-office-dialog-api-with-angular"></a><span data-ttu-id="c936e-130">Angular で Office Dialog API を使用する</span><span class="sxs-lookup"><span data-stu-id="c936e-130">Using the Office Dialog API with Angular</span></span>

<span data-ttu-id="c936e-131">Office のアドインの Dialog API を使えば、アドインでは、メイン ページと情報をやりとりできるセミモーダル ダイアログ ボックスで、ページを開けるようになります。通常、これは作業ウィンドウにあります。</span><span class="sxs-lookup"><span data-stu-id="c936e-131">The Office add-in Dialog API enables your add-in to open a page in a semimodal dialog box that can exchange information with the main page, which is typically in a task pane.</span></span> 

<span data-ttu-id="c936e-p107">[DisplayDialogAsync](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) メソッドは、ダイアログ ボックスで開くべきページの URL を指定するパラメーターを受け取ります。アドインでは、独立した HTML ページ (基本ページとは異なるページ) でこのパラメーターに渡すか、Angular アプリケーションでルートの URL を渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="c936e-p107">The [displayDialogAsync](https://docs.microsoft.com/javascript/api/office/office.ui?view=office-js) method takes a parameter that specifies the URL of the page that should open in the dialog box. Your add-in can have a separate HTML page (different from the base page) to pass to this parameter, or you can pass the URL of a route in your Angular appication.</span></span> 

<span data-ttu-id="c936e-p108">重要な点として、ルートを渡すと、ダイアログ ボックスによって新しいウィンドウとその実行コンテキストが作成されることに注意してください。ダイアログ ボックスで、この新しいコンテキストに対して基本ページとそのすべての初期化、およびブートストラップ コードを再度実行し、すべての変数が初期値に設定されます。この手法により、ダイアログ ボックスで、単一ページのアプリケーションの 2 番目のインスタンスが起動します。ダイアログ ボックス内の変数を変更するコードは、同じ変数の作業ウィンドウのバージョンは変更しません。同様に、ダイアログ ボックスには、それ自体にセッション ストレージがあり、作業ウィンドウからコードでそこにアクセスすることはできません。</span><span class="sxs-lookup"><span data-stu-id="c936e-p108">It is important to remember, if you pass a route, that the dialog box creates a new window with its own execution context. Your base page and all its initialization and bootstrapping code run again in this new context, and any variables are set to their initial values in the dialog box. So this technique launches a second instance of your single page application in the dialog box. Code that changes variables in the dialog box does not change the task pane version of the same variables. Similarly, the dialog box has its own session storage, which is not accessible from code in the task pane.</span></span>  


## <a name="trigger-the-ui-update"></a><span data-ttu-id="c936e-139">UI の更新をトリガーする</span><span class="sxs-lookup"><span data-stu-id="c936e-139">Trigger the UI update</span></span>

<span data-ttu-id="c936e-p109">Angular アプリでは UI が更新されない場合があります。これは、コード部分が Angular ゾーンの外から実行されるためです。解決策としては、次の例に示すように、ゾーン内にコードを配置します。</span><span class="sxs-lookup"><span data-stu-id="c936e-p109">In an Angular app, the UI sometimes does not update. This is because that part of the code runs out of the Angular zone. The solution is to put the code in the zone, as shown in the following example.</span></span>

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

## <a name="using-observable"></a><span data-ttu-id="c936e-143">Observable を使用する</span><span class="sxs-lookup"><span data-stu-id="c936e-143">Using Observable</span></span>

<span data-ttu-id="c936e-p110">Angular は RxJS (JavaScript の事後対応型の拡張機能) を使用し、RxJS は `Observable` と `Observer` のオブジェクトを導入して非同期処理を実装します。このセクションでは、`Observables` の使い方についての概要を簡単に紹介しています。さらに詳細な情報については、[RxJS](http://reactivex.io/rxjs/) の公式ドキュメントを参照してください。</span><span class="sxs-lookup"><span data-stu-id="c936e-p110">Angular uses RxJS (Reactive Extensions for JavaScript), and RxJS introduces `Observable` and `Observer` objects to implement asynchronous processing. This section provides a brief introduction to using `Observables`; for more detailed information, see the official [RxJS](http://reactivex.io/rxjs/) documentation.</span></span>

<span data-ttu-id="c936e-p111">`Observable` は、ある意味で `Promise` オブジェクトに似ています。非同期の呼び出しからすぐに返されますが、すぐには解決されない可能性があります。しかし、`Promise` は、単一の値 (配列オブジェクトのことがあります) なのに対し、`Observable` は、オブジェクトの配列 (メンバーが 1 つだけの可能性あり) です。そのため、コードで `concat`、`map`、`filter` などの[配列メソッド](https://www.w3schools.com/jsref/jsref_obj_array.asp)を `Observable` オブジェクトで呼び出すことができます。</span><span class="sxs-lookup"><span data-stu-id="c936e-p111">An `Observable` is like a `Promise` object in some ways - it is returned immediately from an asynchronous call, but it might not resolve until some time later. However, while a `Promise` is a single value (which can be an array object), an `Observable` is an array of objects (possibly with only a single member). This enables code to call [array methods](https://www.w3schools.com/jsref/jsref_obj_array.asp), such as `concat`, `map`, and `filter`, on `Observable` objects.</span></span> 

### <a name="pushing-instead-of-pulling"></a><span data-ttu-id="c936e-149">プルではなくプッシュ</span><span class="sxs-lookup"><span data-stu-id="c936e-149">Pushing instead of pulling</span></span>

<span data-ttu-id="c936e-p112">コードは `Promise` オブジェクトを変数に割り当てることによって "プル" しますが、`Observable` オブジェクトは、値を `Observable` に*登録*するオブジェクトに、"プッシュ" します。サブスクライバーは、`Observer` オブジェクトです。プッシュ アーキテクチャの利点は、時間の経過と共に新しいメンバーを `Observable` 配列に追加できることです。新しいメンバーが追加されると、`Observable` に登録されるすべての `Observer` オブジェクトは通知を受信します。</span><span class="sxs-lookup"><span data-stu-id="c936e-p112">Your code "pulls" `Promise` objects by assigning them to variables, but `Observable` objects "push" their values to objects that *subscribe* to the `Observable`. The subscribers are `Observer` objects. The benefit of the push architecture is that new members can be added to the `Observable` array over time. When a new member is added, all the `Observer` objects that subscribe to the `Observable` receive a notification.</span></span> 

<span data-ttu-id="c936e-p113">`Observer` は、関数とともに新規の各オブジェクト ("next" オブジェクトと呼ばれる) を処理するように構成されます。(また、エラーと完了の通知に応答するようにも構成されます。例については、次のセクションを参照してください。)このため、`Observable` オブジェクトは、`Promise` オブジェクトよりも幅広いシナリオで使用できます。たとえば、AJAX 呼び出しから `Observable` を返すことに加えて、`Promise` を返し、`Observable` をテキスト ボックスの "変更" イベント ハンドラーなどのイベント ハンドラーから返すことができます。ユーザーがボックスにテキストを入力するたびに、登録されているすべての `Observer` オブジェクトが、最新のテキストや、アプリケーションの現在の状態を入力として使用することによって、すぐに対応します。</span><span class="sxs-lookup"><span data-stu-id="c936e-p113">The `Observer` is configured to process each new object (called the "next" object) with a function. (It is also configured to respond to an error and a completion notification. See the next section for an example.) For this reason, `Observable` objects can be used in a wider range of scenarios than `Promise` objects. For example, in addition to returning an `Observable` from an AJAX call, the way you can return a `Promise`, an `Observable` can be returned from an event handler, such as the "changed" event handler for a text box. Each time a user enters text in the box, all the subscribed `Observer` objects react immediately using the latest text and/or the current state of the application as input.</span></span> 


### <a name="waiting-until-all-asynchronous-calls-have-completed"></a><span data-ttu-id="c936e-159">すべての非同期呼び出しが完了するまで待機する</span><span class="sxs-lookup"><span data-stu-id="c936e-159">Waiting until all asynchronous calls have completed</span></span>

<span data-ttu-id="c936e-160">一連の `Promise` オブジェクトの各メンバーが解決されるときのみ確実にコールバックが実行されるようにしたい場合は、`Promise.all()` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="c936e-160">When you want to ensure that a callback only runs when every member of a set of `Promise` objects has resolved, use the `Promise.all()` method.</span></span>

```js
myPromise.all([x, y, z]).then(
  // TODO: Callback logic goes here
)
``` 

<span data-ttu-id="c936e-161">`Observable` オブジェクトで同じことを行うには、[Observable.forkJoin()](https://github.com/Reactive-Extensions/RxJS/blob/master/doc/api/core/operators/forkjoin.md) メソッドを使います。</span><span class="sxs-lookup"><span data-stu-id="c936e-161">To do the same thing with an `Observable` object, you use the [Observable.forkJoin()](https://github.com/Reactive-Extensions/RxJS/blob/master/doc/api/core/operators/forkjoin.md) method.</span></span>  

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

## <a name="compile-the-angular-application-using-the-ahead-of-time-aot-compiler"></a><span data-ttu-id="c936e-162">Ahead-of-Time (AOT) コンパイラを使って Angular アプリケーションをコンパイルする</span><span class="sxs-lookup"><span data-stu-id="c936e-162">Compile the Angular application using the Ahead-of-Time (AOT) compiler</span></span>

<span data-ttu-id="c936e-163">アプリケーションのパフォーマンスは、ユーザー エクスペリエンスの中でも重要度が高いものの 1 つです。</span><span class="sxs-lookup"><span data-stu-id="c936e-163">Application performance is one of the most important aspects of user experience.</span></span> <span data-ttu-id="c936e-164">Angular アプリケーションは、ビルド時に Angular Ahead-of-Time (AOT) コンパイラを使用してアプリをコンパイルすることで最適化することができます。</span><span class="sxs-lookup"><span data-stu-id="c936e-164">An Angular application can be optimized by using the Angular Ahead-of-Time (AOT) compiler to compile the app at build time.</span></span> <span data-ttu-id="c936e-165">すべてのソース コード (HTML テンプレートと TypeScript) を効率的な JavaScript コードに変換します。</span><span class="sxs-lookup"><span data-stu-id="c936e-165">It converts all source code (HTML templates and TypeScript) into efficient JavaScript code.</span></span> <span data-ttu-id="c936e-166">AOT コンパイラを使用してアプリをコンパイルすると、実行時に追加のコンパイルは実行されません。そのため、HTML テンプレートのレンダリングと非同期要求が高速になります。</span><span class="sxs-lookup"><span data-stu-id="c936e-166">If you compile your app with the AOT compiler, no additional compilation will occur at runtime, which results in faster rendering and faster asynchronous requests for HTML templates.</span></span> <span data-ttu-id="c936e-167">さらに、Angular コンパイラを配布可能なアプリケーションに含める必要がないため、アプリケーション全体のサイズが小さくなります。</span><span class="sxs-lookup"><span data-stu-id="c936e-167">Additionally, the overall application size will be reduced, because the Angular compiler won't need to be included in the application distributable.</span></span> 

<span data-ttu-id="c936e-168">AOT コンパイラを使用するには、`ng build` または `ng serve` コマンドに `--aot` を追加します。</span><span class="sxs-lookup"><span data-stu-id="c936e-168">To use the AOT compiler, add `--aot` to the `ng build` or `ng serve` command:</span></span>

```bash
ng build --aot
ng serve --aot
```

> [!NOTE]
> <span data-ttu-id="c936e-169">Angular Ahead-of-Time (AOT) コンパイラの詳細については、[公式ガイド](https://angular.io/guide/aot-compiler)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c936e-169">To learn more about the Angular Ahead-of-Time (AOT) compiler, see the [official guide](https://angular.io/guide/aot-compiler).</span></span>

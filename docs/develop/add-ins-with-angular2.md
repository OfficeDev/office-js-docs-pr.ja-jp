---
title: Angular で Office アドインを開発する
description: 角度を使用して、単一ページアプリケーションとして Office アドインを作成するためのガイダンスを取得します。
ms.date: 01/27/2020
localization_priority: Normal
ms.openlocfilehash: 2cd90a51f49adfd03c0096d55399012e88da1da0
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608986"
---
# <a name="develop-office-add-ins-with-angular"></a><span data-ttu-id="31886-103">Angular で Office アドインを開発する</span><span class="sxs-lookup"><span data-stu-id="31886-103">Develop Office Add-ins with Angular</span></span>

<span data-ttu-id="31886-104">この記事では、Angular 2+ を使って、単一ページのアプリケーションとして Office アドインを作成する方法を説明します。</span><span class="sxs-lookup"><span data-stu-id="31886-104">This article provides guidance for using Angular 2+ to create an Office Add-in as a single page application.</span></span>

> [!NOTE]
> <span data-ttu-id="31886-p101">Angular を使用して Office アドインを作成した経験を基に、何か投稿する内容がありますか。[GitHub](https://github.com/OfficeDev/office-js-docs) でこの記事に対して投稿するか、リポジトリで[問題](https://github.com/OfficeDev/office-js-docs-pr/issues)を提出することでフィードバックを提出できます。</span><span class="sxs-lookup"><span data-stu-id="31886-p101">Do you have something to contribute based on your experience using Angular to create Office Add-ins? You can contribute to this article in [GitHub](https://github.com/OfficeDev/office-js-docs) or provide your feedback by submitting an [issue](https://github.com/OfficeDev/office-js-docs-pr/issues) in the repo.</span></span> 

<span data-ttu-id="31886-107">Angular フレームワークを使用してビルドされる Office アドインのサンプルについては、「[Angular でビルドする Word スタイル チェック アドイン](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="31886-107">For an Office Add-ins sample that's built using the Angular framework, see [Word Style Checking Add-in Built on Angular](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker).</span></span>

## <a name="install-the-typescript-type-definitions"></a><span data-ttu-id="31886-108">TypeScript 型の定義をインストールする</span><span class="sxs-lookup"><span data-stu-id="31886-108">Install the TypeScript type definitions</span></span>

<span data-ttu-id="31886-109">nodejs ウィンドウを開き、コマンド ラインで次のように入力します:</span><span class="sxs-lookup"><span data-stu-id="31886-109">Open an nodejs window and enter the following at the command line:</span></span>

```command&nbsp;line
npm install --save-dev @types/office-js
```

## <a name="bootstrapping-must-be-inside-officeinitialize"></a><span data-ttu-id="31886-110">ブートス トラップは必ず Office.initialize 内に</span><span class="sxs-lookup"><span data-stu-id="31886-110">Bootstrapping must be inside Office.initialize</span></span>

<span data-ttu-id="31886-111">Office、Word、または Excel の JavaScript Api を呼び出すページでは、まず、コードでプロパティにメソッドを割り当てる必要があり `Office.initialize` ます。</span><span class="sxs-lookup"><span data-stu-id="31886-111">On any page that calls the Office, Word, or Excel JavaScript APIs, your code must first assign a method to the `Office.initialize` property.</span></span> <span data-ttu-id="31886-112">(初期化コードがない場合、メソッドの本体は空の "" 記号になることができ `{}` ますが、このプロパティは未定義のままにしないでください `Office.initialize` 。</span><span class="sxs-lookup"><span data-stu-id="31886-112">(If you have no initialization code, the method body can be just empty "`{}`" symbols, but you must not leave the `Office.initialize` property undefined.</span></span> <span data-ttu-id="31886-113">詳細については、「 [Office アドインを初期化する](initialize-add-in.md)」を参照してください)。Office は、Office JavaScript ライブラリを初期化した直後に、このメソッドを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="31886-113">For details, see [Initialize your Office Add-in](initialize-add-in.md).) Office calls this method immediately after it has initialized the Office JavaScript libraries.</span></span>

<span data-ttu-id="31886-p103">**Angular のブートストラップ コードは `Office.initialize` に割り当てられたメソッドの中で呼び出すことで**、Office の JavaScript ライブラリが最初に初期化されるようにする必要があります。以下は、これを行う方法を示した簡単な例です。このコードは、プロジェクトの main.ts ファイルの中にある必要があります。</span><span class="sxs-lookup"><span data-stu-id="31886-p103">**Your Angular bootstrapping code must be called inside the method that you assign to `Office.initialize`** to ensure that the Office JavaScript libraries have initialized first. The following is a simple example that shows how to do this. This code should be in the main.ts file of the project.</span></span>

```js
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
import { AppModule } from './app.module';

Office.initialize = function () {
  const platform = platformBrowserDynamic();
  platform.bootstrapModule(AppModule);
};
```

## <a name="use-the-hash-location-strategy-in-the-angular-application"></a><span data-ttu-id="31886-117">Angular アプリケーションで Hash Location Strategy を使う</span><span class="sxs-lookup"><span data-stu-id="31886-117">Use the hash location strategy in the Angular application</span></span>

<span data-ttu-id="31886-p104">Hash Location Strategy を指定しないと、アプリケーションでルート間の移動が機能しない可能性があります。2 つの方法のいずれかでこれを行うことができます。1 つ目の方法は、次の例に示すとおり、アプリ モジュールでプロバイダーをロケーションの戦略に指定できます。これは app.module.ts ファイルに入ります。</span><span class="sxs-lookup"><span data-stu-id="31886-p104">Navigating between routes in the application might not work if you don't specify the hash location strategy. You can do this in one of two ways. First, you can specify a provider for the location strategy in your app module, as shown in the following example. It goes into the app.module.ts file.</span></span>

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

<span data-ttu-id="31886-p105">独立したルーティング モジュール内で、ルートを定義する場合は、Hash Location Strategy を指定する別の方法があります。ルーティング モジュールの .ts ファイルで、戦略を指定する `forRoot` 関数に構成オブジェクトを渡します。以下にコードの例を示します。</span><span class="sxs-lookup"><span data-stu-id="31886-p105">If you define your routes in a separate routing module, there is an alternative way to specify the hash location strategy. In your routing module's .ts file, pass a configuration object to the `forRoot` function that specifies the strategy. The following code is an example.</span></span> 

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


## <a name="consider-wrapping-fabric-components-with-angular-components"></a><span data-ttu-id="31886-125">Fabric コンポーネントと Angular コンポーネントとのラッピングについて検討する</span><span class="sxs-lookup"><span data-stu-id="31886-125">Consider wrapping Fabric components with Angular components</span></span>

<span data-ttu-id="31886-126">アドインには [UI Fabric](https://developer.microsoft.com/fabric#) のスタイルを使用することをお勧めしています。</span><span class="sxs-lookup"><span data-stu-id="31886-126">We recommend using [UI Fabric](https://developer.microsoft.com/fabric#) styling in your add-in.</span></span> <span data-ttu-id="31886-127">Web 用の UI Fabric は 2 つの種類で利用可能です。</span><span class="sxs-lookup"><span data-stu-id="31886-127">UI Fabric for the web is available in two flavors:</span></span> 

- <span data-ttu-id="31886-128">[Fabric React](https://developer.microsoft.com/fabric#/controls/web) は、高度にカスタマイズ可能で堅牢な、常に最新版にアップデートされているアクセスしやすいコンポーネントを提供します。</span><span class="sxs-lookup"><span data-stu-id="31886-128">[Fabric React](https://developer.microsoft.com/fabric#/controls/web) provides robust, up-to-date, accessible components that are highly customizable.</span></span>

- <span data-ttu-id="31886-129">[Fabric Core](https://developer.microsoft.com/fabric#/styles/web) は CSS クラスおよび Sass mixin のコレクションで、Fabric の色、アニメーション、フォント、アイコン、グリッドにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="31886-129">[Fabric Core](https://developer.microsoft.com/fabric#/styles/web) is a collection of CSS classes and Sass mixins that give you access to Fabric's colors, animations, fonts, icons and grid.</span></span>

<span data-ttu-id="31886-130">Fabric コンポーネントを、Angular のコンポーネントでラッピングすることによってアドインで使用することを検討してください。</span><span class="sxs-lookup"><span data-stu-id="31886-130">Consider using Fabric components in your add-in by wrapping them in Angular components.</span></span> <span data-ttu-id="31886-131">これを行う方法を説明した例については、「[Angular でビルドする Word スタイル チェック アドイン](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="31886-131">For an example that shows you how to do this, see [Word Style Checking Add-in Built on Angular](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker).</span></span> <span data-ttu-id="31886-132">たとえば、[fabric.textfield.wrapper](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/shared/office-fabric-component-wrappers/fabric.textfield.wrapper.component.ts) で定義されている Angular コンポーネントで Fabric ファイルの TextField.ts をインポートすると、その場所に Fabric コンポーネントが定義されます。</span><span class="sxs-lookup"><span data-stu-id="31886-132">Note, for example, how the Angular component defined in [fabric.textfield.wrapper](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker/blob/master/app/shared/office-fabric-component-wrappers/fabric.textfield.wrapper.component.ts) imports the Fabric file TextField.ts, where the Fabric component is defined.</span></span> 


## <a name="using-the-office-dialog-api-with-angular"></a><span data-ttu-id="31886-133">Angular で Office Dialog API を使用する</span><span class="sxs-lookup"><span data-stu-id="31886-133">Using the Office dialog API with Angular</span></span>

<span data-ttu-id="31886-134">Office のアドインの Dialog API を使えば、アドインでは、メイン ページと情報をやりとりできるセミモードレス ダイアログ ボックスで、ページを開けるようになります。通常、これは作業ウィンドウにあります。</span><span class="sxs-lookup"><span data-stu-id="31886-134">The Office Add-in dialog API enables your add-in to open a page in a nonmodal dialog box that can exchange information with the main page, which is typically in a task pane.</span></span>

<span data-ttu-id="31886-p108">[DisplayDialogAsync](/javascript/api/office/office.ui) メソッドは、ダイアログ ボックスで開くべきページの URL を指定するパラメーターを受け取ります。アドインでは、独立した HTML ページ (基本ページとは異なるページ) でこのパラメーターに渡すか、Angular アプリケーションでルートの URL を渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="31886-p108">The [displayDialogAsync](/javascript/api/office/office.ui) method takes a parameter that specifies the URL of the page that should open in the dialog box. Your add-in can have a separate HTML page (different from the base page) to pass to this parameter, or you can pass the URL of a route in your Angular application.</span></span> 

<span data-ttu-id="31886-p109">重要な点として、ルートを渡すと、ダイアログ ボックスによって新しいウィンドウとその実行コンテキストが作成されることに注意してください。ダイアログ ボックスで、この新しいコンテキストに対して基本ページとそのすべての初期化、およびブートストラップ コードを再度実行し、すべての変数が初期値に設定されます。この手法により、ダイアログ ボックスで、単一ページのアプリケーションの 2 番目のインスタンスが起動します。ダイアログ ボックス内の変数を変更するコードは、同じ変数の作業ウィンドウのバージョンは変更しません。同様に、ダイアログ ボックスには、それ自体にセッション ストレージがあり、作業ウィンドウからコードでそこにアクセスすることはできません。</span><span class="sxs-lookup"><span data-stu-id="31886-p109">It is important to remember, if you pass a route, that the dialog box creates a new window with its own execution context. Your base page and all its initialization and bootstrapping code run again in this new context, and any variables are set to their initial values in the dialog box. So this technique launches a second instance of your single page application in the dialog box. Code that changes variables in the dialog box does not change the task pane version of the same variables. Similarly, the dialog box has its own session storage, which is not accessible from code in the task pane.</span></span>  


## <a name="trigger-the-ui-update"></a><span data-ttu-id="31886-142">UI の更新をトリガーする</span><span class="sxs-lookup"><span data-stu-id="31886-142">Trigger the UI update</span></span>

<span data-ttu-id="31886-p110">Angular アプリでは UI が更新されない場合があります。これは、コード部分が Angular ゾーンの外から実行されるためです。解決策としては、次の例に示すように、ゾーン内にコードを配置します。</span><span class="sxs-lookup"><span data-stu-id="31886-p110">In an Angular app, the UI sometimes does not update. This is because that part of the code runs out of the Angular zone. The solution is to put the code in the zone, as shown in the following example.</span></span>

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

## <a name="using-observable"></a><span data-ttu-id="31886-146">Observable を使用する</span><span class="sxs-lookup"><span data-stu-id="31886-146">Using Observable</span></span>

<span data-ttu-id="31886-p111">Angular は RxJS (JavaScript の事後対応型の拡張機能) を使用し、RxJS は `Observable` と `Observer` のオブジェクトを導入して非同期処理を実装します。このセクションでは、`Observables` の使い方についての概要を簡単に紹介しています。さらに詳細な情報については、[RxJS](https://rxjs-dev.firebaseapp.com/) の公式ドキュメントを参照してください。</span><span class="sxs-lookup"><span data-stu-id="31886-p111">Angular uses RxJS (Reactive Extensions for JavaScript), and RxJS introduces `Observable` and `Observer` objects to implement asynchronous processing. This section provides a brief introduction to using `Observables`; for more detailed information, see the official [RxJS](https://rxjs-dev.firebaseapp.com/) documentation.</span></span>

<span data-ttu-id="31886-p112">`Observable` は、ある意味で `Promise` オブジェクトに似ています。非同期の呼び出しからすぐに返されますが、すぐには解決されない可能性があります。しかし、`Promise` は、単一の値 (配列オブジェクトのことがあります) なのに対し、`Observable` は、オブジェクトの配列 (メンバーが 1 つだけの可能性あり) です。そのため、コードで `concat`、`map`、`filter` などの[配列メソッド](https://www.w3schools.com/jsref/jsref_obj_array.asp)を `Observable` オブジェクトで呼び出すことができます。</span><span class="sxs-lookup"><span data-stu-id="31886-p112">An `Observable` is like a `Promise` object in some ways - it is returned immediately from an asynchronous call, but it might not resolve until some time later. However, while a `Promise` is a single value (which can be an array object), an `Observable` is an array of objects (possibly with only a single member). This enables code to call [array methods](https://www.w3schools.com/jsref/jsref_obj_array.asp), such as `concat`, `map`, and `filter`, on `Observable` objects.</span></span> 

### <a name="pushing-instead-of-pulling"></a><span data-ttu-id="31886-152">プルではなくプッシュ</span><span class="sxs-lookup"><span data-stu-id="31886-152">Pushing instead of pulling</span></span>

<span data-ttu-id="31886-p113">コードは `Promise` オブジェクトを変数に割り当てることによって "プル" しますが、`Observable` オブジェクトは、値を `Observable` に*登録*するオブジェクトに、"プッシュ" します。サブスクライバーは、`Observer` オブジェクトです。プッシュ アーキテクチャの利点は、時間の経過と共に新しいメンバーを `Observable` 配列に追加できることです。新しいメンバーが追加されると、`Observable` に登録されるすべての `Observer` オブジェクトは通知を受信します。</span><span class="sxs-lookup"><span data-stu-id="31886-p113">Your code "pulls" `Promise` objects by assigning them to variables, but `Observable` objects "push" their values to objects that *subscribe* to the `Observable`. The subscribers are `Observer` objects. The benefit of the push architecture is that new members can be added to the `Observable` array over time. When a new member is added, all the `Observer` objects that subscribe to the `Observable` receive a notification.</span></span> 

<span data-ttu-id="31886-p114">`Observer` は、関数とともに新規の各オブジェクト ("next" オブジェクトと呼ばれる) を処理するように構成されます。(また、エラーと完了の通知に応答するようにも構成されます。例については、次のセクションを参照してください。)このため、`Observable` オブジェクトは、`Promise` オブジェクトよりも幅広いシナリオで使用できます。たとえば、AJAX 呼び出しから `Observable` を返すことに加えて、`Promise` を返し、`Observable` をテキスト ボックスの "変更" イベント ハンドラーなどのイベント ハンドラーから返すことができます。ユーザーがボックスにテキストを入力するたびに、登録されているすべての `Observer` オブジェクトが、最新のテキストや、アプリケーションの現在の状態を入力として使用することによって、すぐに対応します。</span><span class="sxs-lookup"><span data-stu-id="31886-p114">The `Observer` is configured to process each new object (called the "next" object) with a function. (It is also configured to respond to an error and a completion notification. See the next section for an example.) For this reason, `Observable` objects can be used in a wider range of scenarios than `Promise` objects. For example, in addition to returning an `Observable` from an AJAX call, the way you can return a `Promise`, an `Observable` can be returned from an event handler, such as the "changed" event handler for a text box. Each time a user enters text in the box, all the subscribed `Observer` objects react immediately using the latest text and/or the current state of the application as input.</span></span> 


### <a name="waiting-until-all-asynchronous-calls-have-completed"></a><span data-ttu-id="31886-162">すべての非同期呼び出しが完了するまで待機する</span><span class="sxs-lookup"><span data-stu-id="31886-162">Waiting until all asynchronous calls have completed</span></span>

<span data-ttu-id="31886-163">一連の `Promise` オブジェクトの各メンバーが解決されるときのみ確実にコールバックが実行されるようにしたい場合は、`Promise.all()` メソッドを使用します。</span><span class="sxs-lookup"><span data-stu-id="31886-163">When you want to ensure that a callback only runs when every member of a set of `Promise` objects has resolved, use the `Promise.all()` method.</span></span>

```js
myPromise.all([x, y, z]).then(
  // TODO: Callback logic goes here
)
``` 

<span data-ttu-id="31886-164">`Observable` オブジェクトで同じことを行うには、[Observable.forkJoin()](https://github.com/Reactive-Extensions/RxJS/blob/master/doc/api/core/operators/forkjoin.md) メソッドを使います。</span><span class="sxs-lookup"><span data-stu-id="31886-164">To do the same thing with an `Observable` object, you use the [Observable.forkJoin()](https://github.com/Reactive-Extensions/RxJS/blob/master/doc/api/core/operators/forkjoin.md) method.</span></span>  

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

## <a name="compile-the-angular-application-using-the-ahead-of-time-aot-compiler"></a><span data-ttu-id="31886-165">Ahead-of-Time (AOT) コンパイラを使って Angular アプリケーションをコンパイルする</span><span class="sxs-lookup"><span data-stu-id="31886-165">Compile the Angular application using the Ahead-of-Time (AOT) compiler</span></span>

<span data-ttu-id="31886-166">アプリケーションのパフォーマンスは、ユーザー エクスペリエンスの中でも重要度が高いものの 1 つです。</span><span class="sxs-lookup"><span data-stu-id="31886-166">Application performance is one of the most important aspects of user experience.</span></span> <span data-ttu-id="31886-167">Angular アプリケーションは、ビルド時に Angular Ahead-of-Time (AOT) コンパイラを使用してアプリをコンパイルすることで最適化することができます。</span><span class="sxs-lookup"><span data-stu-id="31886-167">An Angular application can be optimized by using the Angular Ahead-of-Time (AOT) compiler to compile the app at build time.</span></span> <span data-ttu-id="31886-168">すべてのソース コード (HTML テンプレートと TypeScript) を効率的な JavaScript コードに変換します。</span><span class="sxs-lookup"><span data-stu-id="31886-168">It converts all source code (HTML templates and TypeScript) into efficient JavaScript code.</span></span> <span data-ttu-id="31886-169">AOT コンパイラを使用してアプリをコンパイルすると、実行時に追加のコンパイルは実行されません。そのため、HTML テンプレートのレンダリングと非同期要求が高速になります。</span><span class="sxs-lookup"><span data-stu-id="31886-169">If you compile your app with the AOT compiler, no additional compilation will occur at runtime, which results in faster rendering and faster asynchronous requests for HTML templates.</span></span> <span data-ttu-id="31886-170">さらに、Angular コンパイラを配布可能なアプリケーションに含める必要がないため、アプリケーション全体のサイズが小さくなります。</span><span class="sxs-lookup"><span data-stu-id="31886-170">Additionally, the overall application size will be reduced, because the Angular compiler won't need to be included in the application distributable.</span></span> 

<span data-ttu-id="31886-171">AOT コンパイラを使用するには、`ng build` または `ng serve` コマンドに `--aot` を追加します。</span><span class="sxs-lookup"><span data-stu-id="31886-171">To use the AOT compiler, add `--aot` to the `ng build` or `ng serve` command:</span></span>

```command&nbsp;line
ng build --aot
ng serve --aot
```

> [!NOTE]
> <span data-ttu-id="31886-172">Angular Ahead-of-Time (AOT) コンパイラの詳細については、[公式ガイド](https://angular.io/guide/aot-compiler)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="31886-172">To learn more about the Angular Ahead-of-Time (AOT) compiler, see the [official guide](https://angular.io/guide/aot-compiler).</span></span>

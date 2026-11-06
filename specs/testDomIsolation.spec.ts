/*
 *  Power BI Visualizations
 *
 *  Copyright (c) Microsoft Corporation
 *  All rights reserved.
 *  MIT License
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the ""Software""), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *
 *  The above copyright notice and this permission notice shall be included in
 *  all copies or substantial portions of the Software.
 *
 *  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 *  THE SOFTWARE.
 */

/**
 * testDom (powerbi-visuals-utils-testutils) names each test root "item_" followed by a
 * single random 16 bit number, appends it to the document and returns it by looking that
 * id up. The roots of earlier specs are never removed, so with a few hundred of them two
 * ids end up colliding now and then, and getElementById hands back an older root that
 * already holds the marks of another spec - which fails that spec at random.
 *
 * Once a spec is over its roots are no longer looked up by id, so their id is dropped:
 * a later root can then only ever be found under its own id. The roots stay in the
 * document, so nothing a spec still holds a reference to is detached.
 */
afterEach(() => {
    Array.from(document.querySelectorAll("body > div[id^='item_']"))
        .forEach((element: Element) => element.removeAttribute("id"));
});

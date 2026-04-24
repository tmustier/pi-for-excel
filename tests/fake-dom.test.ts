import assert from "node:assert/strict";

type ConstructorSnapshot = {
  Node?: unknown;
  Element?: unknown;
  HTMLElement?: unknown;
  HTMLInputElement?: unknown;
  HTMLTextAreaElement?: unknown;
  HTMLSelectElement?: unknown;
  CustomEvent?: unknown;
  document?: unknown;
};

class FakeNode extends EventTarget {
  parentElement: FakeElement | null = null;

  get isConnected(): boolean {
    return this.parentElement ? this.parentElement.isConnected : false;
  }

  remove(): void {
    if (!this.parentElement) {
      return;
    }

    const parent = this.parentElement;
    this.parentElement = null;
    parent.removeChild(this);
  }
}

class FakeElement extends FakeNode {
  readonly tagName: string;
  readonly children: FakeElement[] = [];
  readonly dataset: Record<string, string> = {};
  readonly style: Record<string, string> = {};
  readonly classList = {
    toggle: (_token: string, _force?: boolean): boolean => false,
    add: (..._tokens: string[]): void => {},
    remove: (..._tokens: string[]): void => {},
    contains: (_token: string): boolean => false,
  };

  id = "";
  className = "";
  hidden = false;
  textContent = "";
  private readonly attributes = new Map<string, string>();

  constructor(tagName: string) {
    super();
    this.tagName = tagName.toUpperCase();
  }

  override get isConnected(): boolean {
    if (this.parentElement) {
      return this.parentElement.isConnected;
    }

    return false;
  }

  appendChild(child: FakeElement): FakeElement {
    child.remove();
    child.parentElement = this;
    this.children.push(child);
    return child;
  }

  append(...nodes: Array<FakeElement | string>): void {
    for (const node of nodes) {
      if (typeof node === "string") {
        continue;
      }

      this.appendChild(node);
    }
  }

  replaceChildren(...nodes: FakeElement[]): void {
    for (const child of this.children) {
      child.parentElement = null;
    }

    this.children.length = 0;
    for (const node of nodes) {
      this.appendChild(node);
    }
  }

  removeChild(node: FakeNode): void {
    const index = this.children.findIndex((child) => child === node);
    if (index < 0) {
      return;
    }

    this.children[index].parentElement = null;
    this.children.splice(index, 1);
  }

  contains(target: Element): boolean {
    if (!(target instanceof FakeElement)) {
      return false;
    }

    if (target === this) {
      return true;
    }

    for (const child of this.children) {
      if (child.contains(target)) {
        return true;
      }
    }

    return false;
  }

  setAttribute(name: string, value: string): void {
    this.attributes.set(name, value);
    if (name === "id") {
      this.id = value;
    }
    if (name === "class") {
      this.className = value;
    }
  }

  getAttribute(name: string): string | null {
    return this.attributes.get(name) ?? null;
  }

  querySelectorAll<T extends Element = Element>(selector: string): T[] {
    if (selector !== "button") {
      return [];
    }

    const matches: T[] = [];
    const visit = (node: FakeElement) => {
      if (node.tagName === "BUTTON") {
        matches.push(node as unknown as T);
      }

      for (const child of node.children) {
        visit(child);
      }
    };

    for (const child of this.children) {
      visit(child);
    }

    return matches;
  }

  closest<T extends Element = Element>(selector: string): T | null {
    const selectors = selector.split(",").map((part) => part.trim()).filter((part) => part.length > 0);

    for (const part of selectors) {
      if (matchesSelector(this, part)) {
        return this as unknown as T;
      }
    }

    let parent = this.parentElement;
    while (parent) {
      for (const part of selectors) {
        if (matchesSelector(parent, part)) {
          return parent as unknown as T;
        }
      }

      parent = parent.parentElement;
    }

    return null;
  }
}

class FakeHTMLElement extends FakeElement {
  constructor(tagName: string) {
    super(tagName);
  }
}

class FakeBodyElement extends FakeHTMLElement {
  constructor() {
    super("body");
  }

  override get isConnected(): boolean {
    return true;
  }
}

class FakeHTMLInputElement extends FakeHTMLElement {
  type = "text";

  constructor() {
    super("input");
  }
}

class FakeHTMLTextAreaElement extends FakeHTMLElement {
  value = "";

  constructor() {
    super("textarea");
  }
}

class FakeHTMLSelectElement extends FakeHTMLElement {
  constructor() {
    super("select");
  }
}

class FakeCustomEvent extends Event {
  readonly detail: unknown;

  constructor(type: string, init?: CustomEventInit<unknown>) {
    super(type, init);
    this.detail = init?.detail;
  }
}

class FakeDocument extends EventTarget {
  readonly body: FakeHTMLElement;

  constructor() {
    super();
    this.body = new FakeBodyElement();
  }

  createElement(tagName: string): Element {
    const normalized = tagName.trim().toLowerCase();
    if (normalized === "input") {
      return new FakeHTMLInputElement() as unknown as Element;
    }

    if (normalized === "textarea") {
      return new FakeHTMLTextAreaElement() as unknown as Element;
    }

    if (normalized === "select") {
      return new FakeHTMLSelectElement() as unknown as Element;
    }

    return new FakeHTMLElement(normalized) as unknown as Element;
  }

  createElementNS(_namespace: string | null, qualifiedName: string): Element {
    return this.createElement(qualifiedName);
  }

  getElementById(id: string): HTMLElement | null {
    const visit = (node: FakeElement): FakeElement | null => {
      if (node.id === id) {
        return node;
      }

      for (const child of node.children) {
        const match = visit(child);
        if (match) return match;
      }

      return null;
    };

    const found = visit(this.body);
    return found as unknown as HTMLElement | null;
  }

  querySelectorAll(selector: string): Element[] {
    const matches: Element[] = [];

    const visit = (node: FakeElement): void => {
      if (matchesSelector(node, selector)) {
        matches.push(node as unknown as Element);
      }

      for (const child of node.children) {
        visit(child);
      }
    };

    visit(this.body);
    return matches;
  }
}

function matchesSelector(node: FakeElement, selector: string): boolean {
  if (selector === "textarea") {
    return node.tagName === "TEXTAREA";
  }

  if (selector === "input") {
    return node.tagName === "INPUT";
  }

  if (selector === "select") {
    return node.tagName === "SELECT";
  }

  if (selector === "[contenteditable='true']" || selector === "[contenteditable='plaintext-only']") {
    const value = node.getAttribute("contenteditable");
    return value === "true" || value === "plaintext-only";
  }

  if (selector === "[data-claims-escape='true']") {
    return node.dataset.claimsEscape === "true";
  }

  if (selector.startsWith(".")) {
    const cls = selector.slice(1);
    return node.className.split(/\s+/u).includes(cls);
  }

  if (selector.startsWith("#")) {
    return node.id === selector.slice(1);
  }

  return false;
}

export interface FakeDomHandle {
  document: Document;
  restore: () => void;
}

export function installFakeDom(): FakeDomHandle {
  const previous: ConstructorSnapshot = {
    Node: Reflect.get(globalThis, "Node"),
    Element: Reflect.get(globalThis, "Element"),
    HTMLElement: Reflect.get(globalThis, "HTMLElement"),
    HTMLInputElement: Reflect.get(globalThis, "HTMLInputElement"),
    HTMLTextAreaElement: Reflect.get(globalThis, "HTMLTextAreaElement"),
    HTMLSelectElement: Reflect.get(globalThis, "HTMLSelectElement"),
    CustomEvent: Reflect.get(globalThis, "CustomEvent"),
    document: Reflect.get(globalThis, "document"),
  };

  const fakeDocument = new FakeDocument();

  Reflect.set(globalThis, "Node", FakeNode);
  Reflect.set(globalThis, "Element", FakeElement);
  Reflect.set(globalThis, "HTMLElement", FakeHTMLElement);
  Reflect.set(globalThis, "HTMLInputElement", FakeHTMLInputElement);
  Reflect.set(globalThis, "HTMLTextAreaElement", FakeHTMLTextAreaElement);
  Reflect.set(globalThis, "HTMLSelectElement", FakeHTMLSelectElement);
  Reflect.set(globalThis, "CustomEvent", FakeCustomEvent);
  Reflect.set(globalThis, "document", fakeDocument);

  return {
    document: fakeDocument as unknown as Document,
    restore: () => {
      restoreGlobal("Node", previous.Node);
      restoreGlobal("Element", previous.Element);
      restoreGlobal("HTMLElement", previous.HTMLElement);
      restoreGlobal("HTMLInputElement", previous.HTMLInputElement);
      restoreGlobal("HTMLTextAreaElement", previous.HTMLTextAreaElement);
      restoreGlobal("HTMLSelectElement", previous.HTMLSelectElement);
      restoreGlobal("CustomEvent", previous.CustomEvent);
      restoreGlobal("document", previous.document);
    },
  };
}

function restoreGlobal(key: string, value: unknown): void {
  if (value === undefined) {
    const deleted = Reflect.deleteProperty(globalThis, key);
    assert.equal(deleted, true);
    return;
  }

  Reflect.set(globalThis, key, value);
}

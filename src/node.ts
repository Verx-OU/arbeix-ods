import { exit } from "process";
import util from "util";

export type Attrib = Record<string, string>;
export type Item = {
  [P in string]: Item[];
} & {
  [":@"]?: Attrib;
  ["#text"]?: string;
};

function realKey(keys: string[]): string | undefined {
  return keys.filter((j) => ![":@", "#text"].includes(j))[0];
}

export type NodeLike<T> = new (hint: Item, parent?: Node) => T;

export class Node {
  readonly parent?: Node;
  readonly item: Item;
  readonly name: string;
  private _children: Node[];
  public get children(): Node[] {
    return this._children;
  }

  constructor(item: Item, parent?: Node) {
    this.item = item;
    this.name = realKey(Object.keys(item)) ?? "";
    this.parent = parent;
    const items = this.item[this.name] || [];
    this._children = items.map(this.create);
  }

  protected create = (item: Item): Node => {
    const parent = this;
    if (item["#text"] !== undefined) return new TextNode(item, parent);
    return new Node(item, parent);
  };

  get attrib(): Attrib {
    return this.item[":@"] ?? {};
  }

  setAttrib = (key: string, val: string): this => {
    if (this.item[":@"] === undefined) this.item[":@"] = {};
    this.attrib[key] = val;
    return this;
  };

  deleteAttrib = (key: string): this => {
    if (this.item[":@"] === undefined) this.item[":@"] = {};
    delete this.attrib[key];
    return this;
  };

  protected get contained(): Item[] {
    return this.item[this.name]!;
  }

  single = (key: string): Node => this.children.find((i) => i.name == key)!;

  dig = (...keys: string[]): Node => {
    let it = this as Node;
    keys.forEach((key) => (it = it.single(key)));
    return it;
  };

  all = <T extends Node = Node>(
    key: string,
    downcastTo?: NodeLike<T>
  ): NodeSet<T> => {
    const nodes = this.children
      .filter((i) => i.name === key)
      .map((i) => (downcastTo ? i.downcast(downcastTo) : (i as T)));
    return new NodeSet<T>(...nodes);
  };

  walk = (fn: (node: Node) => void) => {
    fn(this);
    this.children.forEach((i) => i.walk(fn));
  };

  downcast = <T>(target: NodeLike<T>): T => {
    return new target(this.item, this.parent);
  };

  clear = (): this => {
    this.contained.length = 0;
    this.children.length = 0;
    return this;
  };

  addNode = (name: string): Node => {
    const newItem = { [name]: [] };
    this.contained.push(newItem);
    const newNode = this.create(newItem);
    this.children.push(newNode);
    return newNode;
  };

  addText = (text: string): TextNode => {
    const newItem = { "#text": text } as Item;
    this.contained.push(newItem);
    const newNode = this.create(newItem) as TextNode;
    this.children.push(newNode);
    return newNode;
  };

  copyContentsFrom = (other: Node): this => {
    const clone = structuredClone(other.item);
    Object.defineProperties(this, {
      parent: { value: other.parent, enumerable: true },
      item: { value: clone, enumerable: true },
      name: { value: realKey(Object.keys(clone))!, enumerable: true },
    });
    this._children = this.item[this.name]!.map(this.create);
    let parentIndex = this.parent!.children.indexOf(this);
    if (parentIndex === -1) throw "This node isn't in its own parent?";
    this.parent!.contained[parentIndex] = clone;
    return this;
  };

  [util.inspect.custom] = (): any => {
    let object: Record<string, any> = { name: this.name };
    if (Object.keys(this.attrib).length > 0)
      object = { ...object, attrib: this.attrib };
    if (this.children.length > 0) object = { ...object, items: this.children };
    return object;
  };
}

export class Document extends Node {
  constructor(root: Item[]) {
    super({ "": root });
  }

  spreadsheet = (): Spreadsheet =>
    this.dig(
      "office:document-content",
      "office:body",
      "office:spreadsheet"
    ).downcast(Spreadsheet);
}

class Spreadsheet extends Node {
  tables = () => this.all("table:table", Table).indexByAttrib("table:name");
}

const FORMULA_REFERENCE = /(?<=[\[:])\.([A-Z]+)(\d+)(?=[\]:])/g;

class Table extends Node {
  forEachRow = (fn: (row: Row) => void) => {
    const rows = this.all("table:table-row", Row);
    let rowIndex = 0;
    rows.forEach((row) => {
      fn(
        Object.defineProperty(row, "rowIndex", {
          value: rowIndex,
          enumerable: true,
        })
      );
      rowIndex += parseInt(row.attrib["table:number-rows-repeated"] ?? "1");
    });
  };

  forEachCell = (fn: (cell: Cell) => void) => {
    this.forEachRow((row) => row.forEachCell(fn));
  };

  // DANGER: doesn't handle inserting between repeated rows
  insertRow = (index: number): Row => {
    let insertBefore = null as Row | null;
    this.forEachRow((row) => {
      if (row.rowIndex === index) insertBefore = row;
    });
    if (insertBefore === null) throw "Invalid index to insert to";
    const actualIndex = this.contained.indexOf(insertBefore.item);
    if (actualIndex === -1) throw "Row to insert at not actually in table?";

    this.forEachCell((c) => c.adjustFormulae(index, +1));

    const newItem: Item = { ["table:table-row"]: [] };
    const newRow: Row = new Row(newItem, this.parent);
    Object.defineProperty(newRow, "rowIndex", {
      value: index,
      enumerable: true,
    });
    this.contained.splice(actualIndex, 0, newItem);
    this.children.splice(actualIndex, 0, newRow);
    return newRow;
  };

  // DANGER: doesn't handle inserting between repeated rows
  deleteRow = (index: number) => {
    let target = null as Row | null;
    this.forEachRow((row) => {
      if (row.rowIndex === index) target = row;
    });
    if (target === null) return;
    let actualIndex = this.contained.indexOf(target.item);
    if (actualIndex === -1) return;

    this.forEachCell((c) => c.adjustFormulae(index, -1));

    this.contained.splice(actualIndex, 1);
    this.children.splice(actualIndex, 1);
  };
}

export class Row extends Node {
  readonly rowIndex!: number;

  forEachCell = (fn: (cell: Cell) => void): this => {
    const cells = this.all("table:table-cell", Cell);
    let colIndex = 0;
    cells.forEach((cell) => {
      fn(
        Object.defineProperties(cell, {
          rowIndex: { value: this.rowIndex, enumerable: true },
          colIndex: { value: colIndex, enumerable: true },
        })
      );
      colIndex += parseInt(cell.attrib["table:number-columns-repeated"] ?? "1");
    });
    return this;
  };

  get table(): Table {
    return this.parent as Table;
  }
}

export class Cell extends Node {
  readonly rowIndex!: number;
  readonly colIndex!: number;

  get row(): Row {
    return this.parent as Row;
  }

  get table(): Table {
    return this.row.table;
  }

  adjustFormulae = (fromIndex: number, adjust: number) => {
    const formula = this.attrib["table:formula"];
    if (formula) {
      const replaced = formula.replaceAll(
        FORMULA_REFERENCE,
        (match, col, row) => {
          if (parseInt(row) - 1 > fromIndex)
            return `.${col}${parseInt(row) + adjust}`;
          return match;
        }
      );
      if (formula === replaced) return;
      this.setFormula(replaced);
    }
  };

  setType = (type: string): this => {
    this.setAttrib("office:value-type", type);
    this.setAttrib("calcext:value-type", type);
    return this;
  };

  setFormula = (formula: string, type?: string): this => {
    if (type) this.setType(type);
    return this.setAttrib("table:formula", formula)
      .deleteAttrib("office:value")
      .deleteAttrib("office:string-value")
      .clear();
  };
}

export class TextNode extends Node {
  get text(): string {
    return this.item["#text"] ?? "";
  }
  set text(val: string | undefined) {
    this.item["#text"] = val;
  }

  [util.inspect.custom] = (): any => util.inspect(this.text, { colors: true });
}

export class NodeSet<T extends Node = Node> extends Array<T> {
  indexByAttrib = (attribKey: string): Record<string, T> => {
    const obj: Record<string, T> = {};
    this.forEach((n) => {
      obj[n.attrib[attribKey]!] = n;
    });
    return obj;
  };
}

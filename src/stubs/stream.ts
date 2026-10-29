export class Readable {
  static from() { return new Readable(); }
  pipe() { return this; }
  on() { return this; }
  read() { return null; }
  destroy() {}
}

export class Writable {
  write() { return true; }
  end() {}
  destroy() {}
}

export class Transform extends Readable {}
export class PassThrough extends Transform {}
export class Duplex extends Readable {}

export default { Readable, Writable, Transform, PassThrough, Duplex };

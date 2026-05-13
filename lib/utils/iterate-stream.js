// Async-generator wrapper around a node Readable stream.
//
// Note: this implementation deliberately does not pause/resume the source.
// Pausing the stream between yields used to be done for backpressure, but
// it interacts badly with unzipper's `Parse` stream (which we consume from
// `WorkbookReader`): when the master zip stream was paused while the
// consumer was busy piping one entry to a temp file, Parse would emit its
// `end` event before all remaining entries had been delivered as `data`
// events. The dropped entries — typically `xl/sharedStrings.xml` and
// `[Content_Types].xml` — manifested downstream as unresolved
// `{ sharedString: <n> }` cell values and `Cannot read properties of
// undefined (reading 'sheets')` crashes in `_parseWorksheet`, depending on
// the exact emission order of a given zip.
//
// All current callers feed bounded-size XML or small zip-entry streams, so
// leaving the source in flowing mode (without explicit backpressure here)
// is safe and matches what the consumers already do via their own queues.
module.exports = async function* iterateStream(stream) {
  const contents = [];
  let waiter = null;
  const wake = () => {
    if (waiter) {
      const resolve = waiter;
      waiter = null;
      resolve();
    }
  };

  stream.on('data', data => {
    contents.push(data);
    wake();
  });

  let ended = false;
  stream.on('end', () => {
    ended = true;
    wake();
  });

  let error = null;
  stream.on('error', err => {
    error = err;
    wake();
  });

  while (true) {
    if (error) throw error;
    if (contents.length > 0) {
      yield contents.shift();
      continue;
    }
    if (ended) return;
    // eslint-disable-next-line no-await-in-loop
    await new Promise(resolve => {
      waiter = resolve;
    });
  }
};

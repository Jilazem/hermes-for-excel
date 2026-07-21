// SSE komut veri yolu: Hermes → Excel paneli.
// Panel /api/commands'a bağlanır (EventSource). Ajan /api/push'a action
// gönderdiğinde, komutlar bağlı tüm panellere yayınlanır.

const clients = new Set(); // her biri: { res, id }
let seq = 0;

export function addClient(res) {
  const id = ++seq;
  const client = { res, id };
  clients.add(client);
  res.writeHead(200, {
    "content-type": "text/event-stream",
    "cache-control": "no-cache, no-transform",
    connection: "keep-alive",
    "x-accel-buffering": "no",
  });
  res.write(`event: ready\ndata: ${JSON.stringify({ id })}\n\n`);
  const ping = setInterval(() => {
    try {
      res.write(`event: ping\ndata: {}\n\n`);
    } catch {
      /* ignore */
    }
  }, 25000);
  res.on("close", () => {
    clearInterval(ping);
    clients.delete(client);
  });
  return client;
}

// actions: normalize edilmiş action listesi. Bağlı panel sayısını döndürür.
export function pushCommands(payload) {
  const data = JSON.stringify(payload);
  let sent = 0;
  for (const c of clients) {
    try {
      c.res.write(`event: commands\ndata: ${data}\n\n`);
      sent++;
    } catch {
      clients.delete(c);
    }
  }
  return sent;
}

export function clientCount() {
  return clients.size;
}

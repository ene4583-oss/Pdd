const headers = {'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff'};
const json = (data, status = 200) => Response.json(data, {status, headers});
const fields = ['price', 'purchase', 'packaging', 'shipping', 'rate', 'other'];
const idPattern = /^[0-9a-f-]{36}$/;
const photoTypes = new Set(['image/webp', 'image/png', 'image/jpeg']);
const maxPhoto = 4 * 1024 * 1024;

function project(object) {
  const {id, hasPhoto, ...product} = JSON.parse(object.customMetadata.product);
  return {...product, id, revision: object.version, photo: hasPhoto ? '/api/products/' + id + '/photo?v=' + object.etag : null};
}

async function route(request, env) {
  const url = new URL(request.url);
  if (request.method === 'GET' && Object.hasOwn(assets, url.pathname)) {
    const asset = assets[url.pathname];
    return new Response(asset.text, {headers: {...headers, 'Content-Type': asset.type}});
  }
  if (!url.pathname.startsWith('/api/')) return json({error: 'Not found'}, 404);
  const owner = request.headers.get('oai-authenticated-user-id');
  if (!owner) return json({error: '请重新登录 ChatGPT。'}, 401);
  if (request.method !== 'GET' && request.headers.get('Origin') !== url.origin) {
    return json({error: '请求来源无效。'}, 403);
  }
  const bucket = env.BUCKET;
  const prefix = 'products/' + encodeURIComponent(owner) + '/';
  if (request.method === 'GET' && url.pathname === '/api/products') {
    const products = [];
    let cursor;
    do {
      const page = await bucket.list({prefix, cursor, include: ['customMetadata']});
      products.push(...page.objects.map(project));
      cursor = page.truncated ? page.cursor : undefined;
    } while (cursor);
    return json(products.sort((a, b) => b.updated - a.updated));
  }

  const match = url.pathname.match(/^\/api\/products\/([^/]+)(\/photo)?$/);
  if (!match || !idPattern.test(match[1])) return json({error: '找不到商品。'}, 404);
  const [, id, photoPath] = match;
  const key = prefix + id;

  if (photoPath && request.method === 'GET') {
    const object = await bucket.get(key);
    if (!object || !JSON.parse(object.customMetadata.product).hasPhoto) return json({error: '找不到照片。'}, 404);
    return new Response(object.body, {headers: {...headers, 'Content-Type': object.httpMetadata.contentType}});
  }
  if (photoPath) return json({error: '请求方式无效。'}, 405);

  if (request.method === 'PUT') {
    if (Number(request.headers.get('Content-Length')) > maxPhoto + 8192) return json({error: '照片文件过大。'}, 413);
    let form, input;
    try {
      form = await request.formData();
      input = JSON.parse(form.get('product'));
    } catch {
      return json({error: '商品资料格式无效。'}, 400);
    }
    if (!input || typeof input.name !== 'string' || !input.name.trim() || input.name.length > 100 ||
        typeof input.sku !== 'string' || input.sku.length > 60 ||
        fields.some(key => !Number.isSafeInteger(input[key]) || input[key] < 0 ||
          input[key] > (key === 'rate' ? 10000 : 99999999900))) {
      return json({error: '请检查商品名称和金额。'}, 400);
    }
    const current = await bucket.get(key);
    if ((current?.version ?? null) !== (input.revision ?? null)) return json({error: '商品已在另一台设备更新，请关闭编辑后刷新。'}, 409);
    const photo = form.get('photo');
    if (photo !== null && (!(photo instanceof File) || !photoTypes.has(photo.type) || photo.size > maxPhoto)) {
      return json({error: '请选择 4 MB 以下的有效照片。'}, 400);
    }
    const keep = Boolean(input.keepPhoto && current && JSON.parse(current.customMetadata.product).hasPhoto);
    const product = {id, name: input.name.trim(), sku: input.sku.trim(), updated: Date.now(), hasPhoto: Boolean(photo || keep)};
    for (const field of fields) product[field] = input[field];
    const result = await bucket.put(key, photo ? await photo.arrayBuffer() : keep ? await current.arrayBuffer() : '', {
      onlyIf: current ? {etagMatches: current.etag} : {etagDoesNotMatch: '*'},
      customMetadata: {product: JSON.stringify(product)},
      httpMetadata: {contentType: photo ? photo.type : keep ? current.httpMetadata.contentType : 'application/octet-stream'},
    });
    if (!result) return json({error: '商品已在另一台设备更新，请刷新后再试。'}, 409);
    return json(project(result));
  }

  if (request.method === 'DELETE') {
    await bucket.delete(key);
    return new Response(null, {status: 204, headers});
  }
  return json({error: '请求方式无效。'}, 405);
}

export default {
  async fetch(request, env) {
    try {
      return await route(request, env);
    } catch (error) {
      console.error(error);
      return json({error: '暂时无法连接云端储存，请稍后重试。'}, 503);
    }
  },
};

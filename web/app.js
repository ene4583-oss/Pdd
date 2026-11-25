const by = id => document.getElementById(id);
const form = by('productForm');
const editor = by('editor');
const amounts = ['price', 'purchase', 'packaging', 'shipping', 'rate', 'other'];
let ready = false, products = [], active, photo, preview, busy = false;
const currencyKey = 'shop-cost-currency';
by('currency').value = localStorage.getItem(currencyKey) || 'SGD';
const currency = () => by('currency').value || 'SGD';
const money = cents => new Intl.NumberFormat('zh-SG', {
  style: 'currency', currency: currency(), currencyDisplay: 'code', maximumFractionDigits: 2, minimumFractionDigits: 2
}).format(cents / 100);
const percentage = value => value == null ? '—' : value.toFixed(1) + '%';

function calculate(p) {
  const fee = Number((BigInt(p.price) * BigInt(p.rate) + 5000n) / 10000n);
  const cost = p.purchase + p.packaging + p.shipping + p.other + fee;
  const profit = p.price - cost;
  return { fee, cost, profit, margin: p.price ? 100 * profit / p.price : null };
}

async function api(path = '', options) {
  const response = await fetch('/api/products' + path, options);
  if (response.status === 204) return;
  const result = await response.json();
  if (!response.ok) throw new Error(result.error || '无法连接云端储存。');
  return result;
}

function element(tag, value, className) {
  const node = document.createElement(tag);
  if (value != null) node.textContent = value;
  if (className) node.className = className;
  return node;
}

function render() {
  const body = by('rows');
  body.replaceChildren();
  const query = by('search').value.trim().toLowerCase();
  const filtered = products.filter(p => p.name.toLowerCase().includes(query) || p.sku.toLowerCase().includes(query));
  for (const p of filtered) {
    const { cost, profit, margin } = calculate(p);
    const row = body.insertRow();
    const info = element('div', null, 'product');
    if (p.photo) {
      const img = element('img');
      img.alt = '';
      img.src = p.photo;
      info.append(img);
    } else {
      const placeholder = element('span', '▧', 'placeholder');
      placeholder.setAttribute('aria-hidden', 'true');
      info.append(placeholder);
    }
    const labels = element('div');
    labels.append(element('strong', p.name));
    if (p.sku) labels.append(element('span', p.sku, 'sku'));
    info.append(labels);
    row.insertCell().append(info);
    row.insertCell().textContent = money(p.price);
    row.insertCell().textContent = money(cost);
    const profitCell = row.insertCell();
    profitCell.textContent = money(profit);
    profitCell.className = profit < 0 ? 'negative' : profit > 0 ? 'positive' : '';
    row.insertCell().textContent = percentage(margin);
    const edit = element('button', '查看 / 编辑');
    edit.setAttribute('aria-label', '查看或编辑 ' + p.name);
    edit.addEventListener('click', () => openEditor(p));
    row.insertCell().append(edit);
  }
  const totals = products.reduce((sum, p) => {
    const result = calculate(p);
    return { cost: sum.cost + result.cost, profit: sum.profit + result.profit };
  }, {cost: 0, profit: 0});
  by('count').textContent = products.length;
  by('avgCost').textContent = products.length ? money(totals.cost / products.length) : '—';
  by('avgProfit').textContent = products.length ? money(totals.profit / products.length) : '—';
  by('empty').hidden = filtered.length > 0;
  by('emptyTitle').textContent = products.length ? '找不到符合条件的商品' : '还没有商品';
  by('emptyText').textContent = products.length ? '试试其他商品名称或 SKU。' : '新增第一个商品，就能看到单件成本和利润。';
  by('emptyAdd').hidden = products.length > 0;
  by('listFooter').textContent = query ? `显示 ${filtered.length} 件，共 ${products.length} 件商品` : `共 ${products.length} 件商品`;
}

function readAmounts() {
  return Object.fromEntries(amounts.map(key => [key, Math.round((Number(form.elements[key].value) || 0) * 100)]));
}

function updateEstimate() {
  const valid = amounts.every(key => form.elements[key].validity.valid || form.elements[key].validity.valueMissing);
  const values = valid ? calculate(readAmounts()) : {};
  for (const id of ['fee', 'cost', 'profit']) by(id).textContent = values[id] == null ? '—' : money(values[id]);
  by('margin').textContent = percentage(values.margin);
}

function showPhoto() {
  if (preview?.startsWith('blob:')) URL.revokeObjectURL(preview);
  preview = photo instanceof Blob ? URL.createObjectURL(photo) : photo || '';
  if (preview) by('photoPreview').src = preview;
  else by('photoPreview').removeAttribute('src');
  by('photoPreview').hidden = !photo;
  by('removePhoto').hidden = !photo;
}

function openEditor(product) {
  if (!ready) return;
  form.reset();
  active = product;
  photo = product?.photo || null;
  by('error').textContent = '';
  by('notice').textContent = '';
  by('editorTitle').textContent = product ? '编辑商品' : '新增商品';
  by('delete').hidden = !product;
  by('formCurrency').textContent = currency();
  if (product) {
    form.elements.name.value = product.name;
    form.elements.sku.value = product.sku;
    for (const key of amounts) form.elements[key].value = product[key] / 100;
  }
  showPhoto();
  updateEstimate();
  editor.showModal();
  form.elements.name.focus();
}

function setBusy(value) {
  busy = value;
  for (const id of ['save', 'delete', 'close', 'cancel', 'photo', 'removePhoto']) by(id).disabled = value;
}

async function refresh() {
  products = await api();
  render();
}

by('add').onclick = by('emptyAdd').onclick = () => openEditor();
by('search').oninput = render;
by('currency').onchange = () => {
  localStorage.setItem(currencyKey, currency());
  render();
};
by('close').onclick = by('cancel').onclick = () => editor.close();
editor.addEventListener('cancel', event => { if (busy) event.preventDefault(); });
editor.addEventListener('close', () => {
  photo = null;
  showPhoto();
});
form.addEventListener('input', updateEstimate);
by('photo').onchange = async () => {
  const file = by('photo').files[0];
  if (!file) return;
  by('error').textContent = '';
  setBusy(true);
  try {
    if (!file.type.startsWith('image/') || file.size > 20 * 1024 * 1024) throw new Error();
    const bitmap = await createImageBitmap(file);
    try {
      const ratio = Math.min(1, 900 / Math.max(bitmap.width, bitmap.height));
      const canvas = document.createElement('canvas');
      canvas.width = Math.max(1, Math.round(bitmap.width * ratio));
      canvas.height = Math.max(1, Math.round(bitmap.height * ratio));
      canvas.getContext('2d').drawImage(bitmap, 0, 0, canvas.width, canvas.height);
      const compressed = await new Promise(resolve => canvas.toBlob(resolve, 'image/webp', .85));
      if (!compressed) throw new Error();
      photo = compressed;
      showPhoto();
    } finally {
      bitmap.close();
    }
  } catch {
    by('photo').value = '';
    by('error').textContent = '无法读取照片。请选择可打开且小于 20 MB 的图片。';
  } finally {
    setBusy(false);
  }
};
by('removePhoto').onclick = () => {
  if (busy) return;
  photo = null;
  by('photo').value = '';
  showPhoto();
};
form.onsubmit = async event => {
  event.preventDefault();
  if (busy) return;
  const name = form.elements.name.value.trim();
  if (!name) {
    by('error').textContent = '请输入商品名称。';
    form.elements.name.focus();
    return;
  }
  const product = {
    id: active?.id || crypto.randomUUID(), name, sku: form.elements.sku.value.trim(),
    ...readAmounts(), revision: active?.revision, keepPhoto: typeof photo === 'string'
  };
  setBusy(true);
  try {
    const data = new FormData();
    data.set('product', JSON.stringify(product));
    if (photo instanceof Blob) data.set('photo', photo, 'product.webp');
    const saved = await api('/' + product.id, {method: 'PUT', body: data});
    products = [saved, ...products.filter(item => item.id !== saved.id)];
    render();
    editor.close();
    by('notice').textContent = '商品已保存。';
  } catch (error) {
    by('error').textContent = error.message || '无法保存商品，请重试。';
  } finally {
    setBusy(false);
  }
};
by('delete').onclick = async () => {
  if (busy || !active || !confirm('确定删除这件商品和照片？')) return;
  setBusy(true);
  try {
    await api('/' + active.id, {method: 'DELETE'});
    products = products.filter(item => item.id !== active.id);
    render();
    editor.close();
    by('notice').textContent = '商品已删除。';
  } catch {
    by('error').textContent = '无法删除商品，请重试。';
  } finally {
    setBusy(false);
  }
};

by('add').disabled = by('emptyAdd').disabled = true;
refresh().then(() => {
  ready = true;
  by('add').disabled = by('emptyAdd').disabled = false;
}).catch(error => {
  by('notice').textContent = error.message || '无法读取商品资料。请重新打开页面。';
});
window.addEventListener('focus', () => {
  if (ready && !editor.open && !busy) refresh().catch(() => {});
});

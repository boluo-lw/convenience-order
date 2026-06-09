const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const path = require('path');
const multer = require('multer');
const fs = require('fs');
const QRCode = require('qrcode');
const XLSX = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;

// 存储 SSE 连接
let sseConnections = [];

// 确保uploads目录存在
const uploadsDir = path.join(__dirname, 'public', 'uploads');
if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir, { recursive: true });
}

// 配置multer文件上传
const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, uploadsDir);
    },
    filename: function (req, file, cb) {
        const ext = path.extname(file.originalname);
        const filename = Date.now() + '-' + Math.random().toString(36).substring(2, 9) + ext;
        cb(null, filename);
    }
});

const upload = multer({
    storage: storage,
    limits: {
        fileSize: 5 * 1024 * 1024 // 5MB
    },
    fileFilter: function (req, file, cb) {
        const allowedTypes = ['image/jpeg', 'image/jpg', 'image/png', 
                             'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
                             'application/vnd.ms-excel']; // .xls
        if (allowedTypes.includes(file.mimetype)) {
            cb(null, true);
        } else {
            cb(new Error('只支持jpg、png格式的图片和Excel文件'));
        }
    }
});

// Vercel serverless 环境下使用内存存储（演示目的）
let database = {
  orders: [],
  products: [
    { id: 1, name: '西瓜果切（小份）', price: 8, category: '水果', image: '' },
    { id: 2, name: '西瓜果切（大份）', price: 15, category: '水果', image: '' },
    { id: 3, name: '矿泉水', price: 2, category: '饮料', image: '' },
    { id: 4, name: '可乐', price: 3, category: '饮料', image: '' },
    { id: 5, name: '雪碧', price: 3, category: '饮料', image: '' },
    { id: 6, name: '薯片', price: 5, category: '零食', image: '' },
    { id: 7, name: '饼干', price: 4, category: '零食', image: '' },
    { id: 8, name: '巧克力', price: 6, category: '零食', image: '' }
  ],
  companies: ['A公司', 'B公司', 'C公司', 'D公司', 'E公司'],
  settings: {
    minOrderAmount: 20, // 起送价格，默认20元
    shopAddress: '智谷科技园A栋1楼', // 店铺地址
    phone: '13800138000', // 联系方式
    businessHours: '09:00 - 21:00', // 营业时间段
    adminPassword: 'admin123', // 管理员密码
    wechatQRCode: '', // 微信收款二维码
    alipayQRCode: '' // 支付宝收款二维码
  }
};

// 模拟 lowdb 的操作
const db = {
  get: function(key) {
    return {
      value: function() { return database[key] || []; },
      push: function(item) { 
        database[key].push(item);
        return { write: () => {} };
      },
      find: function(query) { 
        return {
          value: function() { 
            return database[key].find(item => Object.keys(query).every(k => item[k] === query[k])) || null;
          },
          assign: function(update) { 
            const item = database[key].find(i => Object.keys(query).every(k => i[k] === query[k]));
            if (item) Object.assign(item, update);
            return { write: () => {} };
          },
          remove: function(query) { 
            database[key] = database[key].filter(item => !Object.keys(query).every(k => item[k] === query[k]));
            return { write: () => {} };
          }
        };
      },
      write: function() {}
    };
  }
};

// 管理员密码（实际部署时应该使用环境变量或加密存储）
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || 'admin123';

// 中间件
app.use(cors());
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static('public'));

// ============ 图片上传 API ============

// 单文件上传
app.post('/api/upload', upload.single('image'), (req, res) => {
    if (!req.file) {
        return res.status(400).json({ success: false, message: '请选择要上传的图片' });
    }
    
    const imageUrl = `/uploads/${req.file.filename}`;
    res.json({ 
        success: true, 
        message: '上传成功',
        url: imageUrl 
    });
});

// 处理multer错误
app.use(function (err, req, res, next) {
    if (err instanceof multer.MulterError) {
        if (err.code === 'LIMIT_FILE_SIZE') {
            return res.status(400).json({ success: false, message: '图片大小不能超过2MB' });
        }
        return res.status(400).json({ success: false, message: '上传失败: ' + err.message });
    } else if (err) {
        return res.status(400).json({ success: false, message: err.message });
    }
    next();
});

// ============ 用户端 API ============

// 获取商品列表
app.get('/api/products', (req, res) => {
  const products = db.get('products').value();
  res.json({ success: true, data: products });
});

// 获取公司列表
app.get('/api/companies', (req, res) => {
  const companies = db.get('companies').value();
  res.json({ success: true, data: companies });
});

// 获取设置（起送价格）
app.get('/api/settings', (req, res) => {
  const settings = db.get('settings').value();
  res.json({ success: true, data: settings });
});

// 提交订单
app.post('/api/orders', (req, res) => {
  const { company, name, phone, address, items, remark } = req.body;
  
  if (!company || !name || !phone || !address || !items || items.length === 0) {
    return res.status(400).json({ 
      success: false, 
      message: '请填写完整信息并选择商品' 
    });
  }
  
  const products = db.get('products').value();
  const settings = db.get('settings').value();
  let totalPrice = 0;
  const orderItems = items.map(item => {
    const product = products.find(p => p.id === item.productId);
    const subtotal = product.price * item.quantity;
    totalPrice += subtotal;
    return {
      productId: item.productId,
      productName: product.name,
      price: product.price,
      quantity: item.quantity,
      subtotal: subtotal
    };
  });
  
  // 验证起送价格
  if (totalPrice < settings.minOrderAmount) {
    return res.status(400).json({ 
      success: false, 
      message: `订单金额不足起送价（¥${settings.minOrderAmount}），当前金额为 ¥${totalPrice}` 
    });
  }
  
  const order = {
    id: Date.now().toString(),
    orderNo: `ORD${Date.now()}`,
    company,
    name,
    phone,
    address,
    items: orderItems,
    totalPrice,
    remark: remark || '',
    status: '待备货',
    createTime: new Date().toISOString(),
    updateTime: new Date().toISOString()
  };
  
  db.get('orders').push(order).write();
  
  // 通知所有管理端有新订单
  notifyNewOrder(order);
  
  res.json({ 
    success: true, 
    message: '下单成功',
    data: { orderNo: order.orderNo }
  });
});

// 确认支付
app.put('/api/orders/pay', (req, res) => {
  const { orderNo } = req.body;
  
  if (!orderNo) {
    return res.status(400).json({ success: false, message: '缺少订单号' });
  }
  
  const orders = db.get('orders').value();
  const order = orders.find(o => o.orderNo === orderNo);
  
  if (!order) {
    return res.status(404).json({ success: false, message: '订单不存在' });
  }
  
  if (order.status === '已完成') {
    return res.status(400).json({ success: false, message: '订单已完成' });
  }
  
  // 更新订单状态为已支付
  order.status = '已支付';
  order.updateTime = new Date().toISOString();
  
  // 通知管理端订单状态更新
  notifyNewOrder(order);
  
  res.json({ success: true, message: '支付确认成功', data: order });
});

// 生成微信支付二维码
app.get('/api/payment/wechat', async (req, res) => {
  try {
    const { amount, orderNo } = req.query;
    
    if (!amount || !orderNo) {
      return res.status(400).json({ success: false, message: '缺少参数' });
    }
    
    // 生成微信支付链接（模拟，实际使用时需要替换为真实的微信收款码链接）
    // 格式说明：微信个人收款码不能直接指定金额，这里生成包含订单信息的二维码
    const paymentInfo = `微信支付:¥${amount}|订单号:${orderNo}`;
    
    // 生成二维码图片（Base64格式）
    const qrCodeBase64 = await QRCode.toDataURL(paymentInfo, {
      width: 200,
      margin: 2
    });
    
    res.json({
      success: true,
      data: {
        qrCode: qrCodeBase64,
        amount: amount,
        orderNo: orderNo,
        paymentInfo: paymentInfo
      }
    });
  } catch (error) {
    console.error('生成微信支付二维码失败:', error);
    res.status(500).json({ success: false, message: '生成二维码失败' });
  }
});

// 生成支付宝支付二维码
app.get('/api/payment/alipay', async (req, res) => {
  try {
    const { amount, orderNo } = req.query;
    
    if (!amount || !orderNo) {
      return res.status(400).json({ success: false, message: '缺少参数' });
    }
    
    // 生成支付宝支付链接
    // 支付宝个人收款码格式：https://qr.alipay.com/fkxXXXXX
    // 这里生成包含订单信息的二维码
    const paymentInfo = `支付宝支付:¥${amount}|订单号:${orderNo}`;
    
    // 生成二维码图片（Base64格式）
    const qrCodeBase64 = await QRCode.toDataURL(paymentInfo, {
      width: 200,
      margin: 2
    });
    
    res.json({
      success: true,
      data: {
        qrCode: qrCodeBase64,
        amount: amount,
        orderNo: orderNo,
        paymentInfo: paymentInfo
      }
    });
  } catch (error) {
    console.error('生成支付宝支付二维码失败:', error);
    res.status(500).json({ success: false, message: '生成二维码失败' });
  }
});

// SSE 端点 - 管理端订阅新订单通知
app.get('/api/admin/order-notification', (req, res) => {
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  res.setHeader('Access-Control-Allow-Origin', '*');
  
  // 将连接添加到列表
  sseConnections.push(res);
  
  // 发送初始连接成功消息
  res.write('event: connected\ndata: {"message": "已连接到订单通知服务"}\n\n');
  
  // 连接关闭时移除
  req.on('close', () => {
    const index = sseConnections.indexOf(res);
    if (index !== -1) {
      sseConnections.splice(index, 1);
    }
  });
});

// 通知新订单
function notifyNewOrder(order) {
  const eventData = JSON.stringify({
    type: 'new_order',
    data: {
      orderNo: order.orderNo,
      orderId: order.id,
      createTime: order.createTime,
      totalPrice: order.totalPrice
    }
  });
  
  sseConnections.forEach(conn => {
    try {
      conn.write(`event: new_order\ndata: ${eventData}\n\n`);
    } catch (error) {
      // 移除无效连接
      const index = sseConnections.indexOf(conn);
      if (index !== -1) {
        sseConnections.splice(index, 1);
      }
    }
  });
}

// ============ 管理端 API ============

// 管理员登录
app.post('/api/admin/login', (req, res) => {
  const { password } = req.body;
  const settings = db.get('settings').value();
  const adminPassword = settings.adminPassword || ADMIN_PASSWORD;
  
  if (password === adminPassword) {
    res.json({ success: true, message: '登录成功' });
  } else {
    res.status(401).json({ success: false, message: '密码错误' });
  }
});

// 获取设置（起送价格）- 管理端
app.get('/api/admin/settings', (req, res) => {
  const settings = db.get('settings').value();
  res.json({ success: true, data: settings });
});

// 更新设置（起送价格、店铺信息）- 管理端
app.put('/api/admin/settings', (req, res) => {
  const { minOrderAmount, shopAddress, phone, businessHours } = req.body;
  
  const settings = db.get('settings').value();
  
  // 更新起送价格（验证）
  if (minOrderAmount !== undefined) {
    if (isNaN(minOrderAmount) || minOrderAmount < 0) {
      return res.status(400).json({ success: false, message: '起送价格必须是大于等于0的数字' });
    }
    settings.minOrderAmount = parseFloat(minOrderAmount);
  }
  
  // 更新店铺地址
  if (shopAddress !== undefined) {
    settings.shopAddress = shopAddress;
  }
  
  // 更新联系方式
  if (phone !== undefined) {
    settings.phone = phone;
  }
  
  // 更新营业时间段
  if (businessHours !== undefined) {
    settings.businessHours = businessHours;
  }
  
  res.json({ success: true, message: '设置更新成功', data: settings });
});

// 上传收款二维码
app.post('/api/admin/upload-qr', upload.single('qrFile'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, message: '请选择要上传的图片' });
    }
    
    const type = req.body.type;
    if (!type || (type !== 'wechat' && type !== 'alipay')) {
      return res.status(400).json({ success: false, message: '无效的二维码类型' });
    }
    
    const settings = db.get('settings').value();
    const qrCodeUrl = `/uploads/${req.file.filename}`;
    
    if (type === 'wechat') {
      settings.wechatQRCode = qrCodeUrl;
    } else {
      settings.alipayQRCode = qrCodeUrl;
    }
    
    res.json({ 
      success: true, 
      message: '二维码上传成功', 
      data: { url: qrCodeUrl } 
    });
  } catch (error) {
    console.error('上传二维码失败:', error);
    res.status(500).json({ success: false, message: '上传失败，请重试' });
  }
});

// 删除收款二维码
app.delete('/api/admin/remove-qr', (req, res) => {
  const { type } = req.body;
  
  if (!type || (type !== 'wechat' && type !== 'alipay')) {
    return res.status(400).json({ success: false, message: '无效的二维码类型' });
  }
  
  const settings = db.get('settings').value();
  
  if (type === 'wechat') {
    settings.wechatQRCode = '';
  } else {
    settings.alipayQRCode = '';
  }
  
  res.json({ success: true, message: '二维码删除成功' });
});

// 获取支付二维码
app.get('/api/payment/qr-code', (req, res) => {
  const { type } = req.query;
  
  if (!type || (type !== 'wechat' && type !== 'alipay')) {
    return res.status(400).json({ success: false, message: '无效的二维码类型' });
  }
  
  const settings = db.get('settings').value();
  const qrCodeUrl = type === 'wechat' ? settings.wechatQRCode : settings.alipayQRCode;
  
  if (!qrCodeUrl) {
    return res.status(404).json({ success: false, message: '未设置收款二维码' });
  }
  
  res.json({ 
    success: true, 
    data: { qrCodeUrl } 
  });
});

// 下载 Excel 模板
app.get('/api/admin/download-template', (req, res) => {
  // 创建模板工作簿
  const workbook = XLSX.utils.book_new();
  
  // 创建模板数据
  const templateData = [
    ['商品名称', '价格', '分类', '图片 URL'],
    ['示例商品 1', '10.00', '零食', ''],
    ['示例商品 2', '15.00', '饮料', ''],
    ['示例商品 3', '8.00', '水果', '']
  ];
  
  // 创建工作表
  const worksheet = XLSX.utils.aoa_to_sheet(templateData);
  
  // 设置列宽
  worksheet['!cols'] = [
    { wch: 20 },  // 商品名称
    { wch: 10 },  // 价格
    { wch: 10 },  // 分类
    { wch: 30 }   // 图片 URL
  ];
  
  // 添加工作表到工作簿
  XLSX.utils.book_append_sheet(workbook, worksheet, '商品模板');
  
  // 生成 Excel 文件
  const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });
  
  // 设置响应头（使用英文文件名避免编码问题）
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="product_template.xlsx"');
  
  // 返回文件
  res.send(excelBuffer);
});

// 导入商品Excel
app.post('/api/admin/import-products', upload.single('excelFile'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, message: '请选择要导入的Excel文件' });
    }
    
    // 读取Excel文件
    const filePath = path.join(__dirname, 'public', 'uploads', req.file.filename);
    const workbook = XLSX.readFile(filePath);
    
    // 获取第一个工作表
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    // 转换为JSON数据
    const jsonData = XLSX.utils.sheet_to_json(worksheet, {
      header: ['name', 'price', 'category', 'image'],
      skipHeader: true
    });
    
    // 获取现有商品列表
    const products = db.get('products').value();
    
    let successCount = 0;
    let errorCount = 0;
    
    // 遍历数据并添加商品
    jsonData.forEach((row, index) => {
      try {
        // 验证必填字段
        if (!row.name || !row.price) {
          errorCount++;
          return;
        }
        
        // 检查商品是否已存在
        const exists = products.some(p => p.name === row.name);
        if (exists) {
          errorCount++;
          return;
        }
        
        // 创建商品
        const newProduct = {
          id: Date.now() + index,
          name: row.name,
          price: parseFloat(row.price) || 0,
          category: row.category || '其他',
          image: row.image || ''
        };
        
        // 添加到数据库
        products.push(newProduct);
        successCount++;
      } catch (error) {
        errorCount++;
      }
    });
    
    // 删除临时文件
    fs.unlinkSync(filePath);
    
    res.json({
      success: true,
      message: `导入完成，成功 ${successCount} 条，失败 ${errorCount} 条`,
      data: { successCount, errorCount }
    });
  } catch (error) {
    console.error('导入商品失败:', error);
    res.status(500).json({ success: false, message: '导入失败，请检查Excel文件格式' });
  }
});

// 修改管理员密码
app.put('/api/admin/change-password', (req, res) => {
  const { oldPassword, newPassword } = req.body;
  
  if (!oldPassword || !newPassword) {
    return res.status(400).json({ success: false, message: '请输入旧密码和新密码' });
  }
  
  const settings = db.get('settings').value();
  const currentPassword = settings.adminPassword || ADMIN_PASSWORD;
  
  if (oldPassword !== currentPassword) {
    return res.status(400).json({ success: false, message: '旧密码错误' });
  }
  
  if (newPassword.length < 6) {
    return res.status(400).json({ success: false, message: '新密码长度不能少于6位' });
  }
  
  settings.adminPassword = newPassword;
  
  res.json({ success: true, message: '密码修改成功，请使用新密码重新登录' });
});

// 获取所有订单
app.get('/api/admin/orders', (req, res) => {
  const { company, status, startDate, endDate } = req.query;
  
  let orders = db.get('orders').value();
  
  if (company && company !== 'all') {
    orders = orders.filter(o => o.company === company);
  }
  
  if (status && status !== 'all') {
    orders = orders.filter(o => o.status === status);
  }
  
  if (startDate) {
    const start = new Date(startDate);
    orders = orders.filter(o => new Date(o.createTime) >= start);
  }
  
  if (endDate) {
    const end = new Date(endDate);
    end.setHours(23, 59, 59, 999);
    orders = orders.filter(o => new Date(o.createTime) <= end);
  }
  
  orders.sort((a, b) => new Date(b.createTime) - new Date(a.createTime));
  
  res.json({ success: true, data: orders });
});

// 更新订单状态
app.put('/api/admin/orders/:id/status', (req, res) => {
  const { id } = req.params;
  const { status } = req.body;
  
  const validStatuses = ['待备货', '已备货', '已配送', '已完成'];
  if (!validStatuses.includes(status)) {
    return res.status(400).json({ 
      success: false, 
      message: '无效的订单状态' 
    });
  }
  
  const order = db.get('orders').find({ id }).value();
  
  if (!order) {
    return res.status(404).json({ 
      success: false, 
      message: '订单不存在' 
    });
  }
  
  db.get('orders')
    .find({ id })
    .assign({ 
      status, 
      updateTime: new Date().toISOString() 
    })
    .write();
  
  res.json({ 
    success: true, 
    message: '状态更新成功' 
  });
});

// 删除订单
app.delete('/api/admin/orders/:id', (req, res) => {
  const { id } = req.params;
  
  const order = db.get('orders').find({ id }).value();
  
  if (!order) {
    return res.status(404).json({ 
      success: false, 
      message: '订单不存在' 
    });
  }
  
  db.get('orders').remove({ id }).write();
  
  res.json({ 
    success: true, 
    message: '订单删除成功' 
  });
});

// 导出订单（Excel格式）
app.get('/api/admin/orders/export', (req, res) => {
  const { company, status, startDate, endDate } = req.query;
  
  let orders = db.get('orders').value();
  
  if (company && company !== 'all') {
    orders = orders.filter(o => o.company === company);
  }
  
  if (status && status !== 'all') {
    orders = orders.filter(o => o.status === status);
  }
  
  if (startDate) {
    const start = new Date(startDate);
    orders = orders.filter(o => new Date(o.createTime) >= start);
  }
  
  if (endDate) {
    const end = new Date(endDate);
    end.setHours(23, 59, 59, 999);
    orders = orders.filter(o => new Date(o.createTime) <= end);
  }
  
  let csv = '\ufeff';
  csv += '订单号,下单时间,公司,姓名,电话,地址,商品明细,数量,总价,状态,备注\n';
  
  orders.forEach(order => {
    const itemsDetail = order.items.map(i => `${i.productName}(${i.quantity}份)`).join('; ');
    const totalQuantity = order.items.reduce((sum, i) => sum + i.quantity, 0);
    
    csv += `${order.orderNo},`;
    csv += `${new Date(order.createTime).toLocaleString('zh-CN')},`;
    csv += `${order.company},`;
    csv += `${order.name},`;
    csv += `${order.phone},`;
    csv += `${order.address},`;
    csv += `"${itemsDetail}",`;
    csv += `${totalQuantity},`;
    csv += `${order.totalPrice}元,`;
    csv += `${order.status},`;
    csv += `${order.remark}\n`;
  });
  
  res.setHeader('Content-Type', 'text/csv; charset=utf-8');
  res.setHeader('Content-Disposition', 'attachment; filename=orders.csv');
  res.send(csv);
});

// ============ 商品管理 API ============

// 添加商品
app.post('/api/admin/products', (req, res) => {
  const { name, price, category, image } = req.body;
  
  if (!name || !price) {
    return res.status(400).json({ 
      success: false, 
      message: '商品名称和价格不能为空' 
    });
  }
  
  const products = db.get('products').value();
  const newId = products.length > 0 ? Math.max(...products.map(p => p.id)) + 1 : 1;
  
  const newProduct = {
    id: newId,
    name,
    price: parseFloat(price),
    category: category || '其他',
    image: image || ''
  };
  
  db.get('products').push(newProduct).write();
  
  res.json({ 
    success: true, 
    message: '商品添加成功',
    data: newProduct
  });
});

// 更新商品
app.put('/api/admin/products/:id', (req, res) => {
  const { id } = req.params;
  const { name, price, category, image } = req.body;
  
  const product = db.get('products').find({ id: parseInt(id) }).value();
  
  if (!product) {
    return res.status(404).json({ 
      success: false, 
      message: '商品不存在' 
    });
  }
  
  const updateData = {
    name: name || product.name,
    price: price ? parseFloat(price) : product.price,
    category: category || product.category
  };
  
  if (image !== undefined) {
    updateData.image = image;
  }
  
  db.get('products')
    .find({ id: parseInt(id) })
    .assign(updateData)
    .write();
  
  res.json({ 
    success: true, 
    message: '商品更新成功' 
  });
});

// 删除商品
app.delete('/api/admin/products/:id', (req, res) => {
  const { id } = req.params;
  
  const product = db.get('products').find({ id: parseInt(id) }).value();
  
  if (!product) {
    return res.status(404).json({ 
      success: false, 
      message: '商品不存在' 
    });
  }
  
  db.get('products').remove({ id: parseInt(id) }).write();
  
  res.json({ 
    success: true, 
    message: '商品删除成功' 
  });
});

// Vercel serverless 兼容
if (process.env.VERCEL) {
  module.exports = app;
} else {
  app.listen(PORT, () => {
    console.log(`\n========================================`);
    console.log(`智谷超市管理系统已启动！`);
    console.log(`========================================`);
    console.log(`用户端: http://localhost:${PORT}`);
    console.log(`管理端: http://localhost:${PORT}/admin.html`);
    console.log(`管理员密码: ${ADMIN_PASSWORD}`);
    console.log(`========================================\n`);
  });
}
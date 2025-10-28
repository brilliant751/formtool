const ExcelJS = require('../../utils/exceljs.min.js');

Page({

  /**
   * 页面的初始数据
   */
  data: {
    formData: {
      number: '',
      location: '',
      time: '',
      type: '',
      pipeNumber: '',
      img: ''
    },
    // 错误提示
    errors: {
      number: '',
      location: '',
      time: '',
      type: '',
      pipeNumber: '',
      img: ''
    },
    reflection: {
      number: '船号',
      location: '位置',
      time: '时间',
      type: '类型',
      pipeNumber: '管件号',
      img: '图片'
    },
    locationList: ["管加车间", "低温车间", "特种车间", "镀锌车间", "涂装车间", "生管交接"],
    typeList: ["缺失", "变形", "错管"]
  },

  /**
   * 生命周期函数--监听页面加载
   */
  onLoad(options) {

  },

  /**
   * 生命周期函数--监听页面初次渲染完成
   */
  onReady() {

  },

  /**
   * 生命周期函数--监听页面显示
   */
  onShow() {

  },

  /**
   * 生命周期函数--监听页面隐藏
   */
  onHide() {

  },

  /**
   * 生命周期函数--监听页面卸载
   */
  onUnload() {

  },

  /**
   * 页面相关事件处理函数--监听用户下拉动作
   */
  onPullDownRefresh() {

  },

  /**
   * 页面上拉触底事件的处理函数
   */
  onReachBottom() {

  },

  /**
   * 用户点击右上角分享
   */
  onShareAppMessage() {

  },

  onInputChange(e) {
    console.log(e);
    console.log(e.detail);
    const field = e.currentTarget.dataset.field;
    console.log(field);
    const value = e.detail.value;
    this.setData({
      [`formData.${field}`]: value,
      [`errors.${field}`]: '' // 清除错误提示
    });
  },

  onTimeChange(e) {
    console.log(e.detail);
    const value = e.detail.value;
    this.setData({
      'formData.time': value,
      'errors.time': '' // 清除错误提示
    });
  },

  onLocationChange(e) {
    console.log(e.detail);
    const value = e.detail.value;
    this.setData({
      'formData.location': this.data.locationList[value],
      'errors.location': '' // 清除错误提示
    });
  },

  onTypeChange(e) {
    console.log(e.detail);
    const value = e.detail.value;
    this.setData({
      'formData.type': this.data.typeList[value],
      'errors.type': '' // 清除错误提示
    });
  },



  onUploadImage(e) {
    console.log("上传图片");
    wx.chooseImage({
      count: 1,
      sourceType: ['album', 'camera'],
      success: (res) => {
        console.log(res.tempFilePaths);
        this.setData({
          'formData.img': res.tempFilePaths[0],
          'errors.img': '' // 清除错误提示
        });
        console.log(this.data.formData.img);
      }
    })
  },

  onDeleteImage(e) {
    console.log("删除图片");
    this.setData({
      'formData.img': ''
    });
  },

  fetchData() {
    const data = this.data.formData;
    console.log(data);
    const reflection = this.data.reflection;
    const res = {};
    for (const key in data) {
      res[reflection[key]] = key == 'img' ? '' : data[key];
    }
    console.log(res);
    return [res];
  },

  // 验证表单数据
  checkData() {
    const { formData, reflection } = this.data;
    let isValid = true;
    const errors = {
      number: '',
      location: '',
      time: '',
      type: '',
      pipeNumber: '',
      img: ''
    };

    // 检查每个必填字段
    const requiredFields = ['number', 'location', 'time', 'type', 'pipeNumber'];
    
    requiredFields.forEach(field => {
      if (!formData[field] || formData[field].trim() === '') {
        errors[field] = `请输入${reflection[field]}`;
        isValid = false;
      }
    });

    // 更新错误提示
    this.setData({ errors });

    if (!isValid) {
      wx.showToast({
        title: '请填写完整信息',
        icon: 'none',
        duration: 2000
      });
    }

    return isValid;
  },

  async onSubmit() {
    if (!this.checkData()) return;
    console.log("表单提交中。。。");
    try {
      const fs = wx.getFileSystemManager();
      const filepath = `${wx.env.USER_DATA_PATH}/test.xlsx`;
      const bookdata = await this.generateXLSX(this.fetchData());
      
      // 将 ArrayBuffer 转换为 Uint8Array
      const uint8Array = new Uint8Array(bookdata);
      
      fs.writeFile({
        filePath: filepath,
        data: uint8Array.buffer, // 使用 buffer 属性
        success: () => {
          console.log("文件保存成功");
          wx.openDocument({
            filePath: filepath,
            fileType: 'xlsx',
            showMenu: true
          });
        },
        fail: (err) => {
          console.error("文件保存失败", err);
        }
      });
    } catch (err) {
      console.error("提交失败", err);
      wx.showToast({
        title: '导出失败',
        icon: 'none'
      });
    }
  },

  // 读取图片二进制数据
  readImageFile(path) {
    return new Promise((resolve, reject) => {
      const fs = wx.getFileSystemManager();
      fs.readFile({
        filePath: path,
        encoding: 'base64', // 以base64格式读取
        success(res) {
          // 返回base64格式的图片数据
          resolve(res.data);
        },
        fail(err) {
          reject(err);
        }
      });
    });
  },

  // 使用 ExcelJS 生成 Excel 文件
  async generateXLSX(data) {
    try {
      // 创建工作簿
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Sheet1');

      // 获取数据的键（表头）
      const headers = Object.keys(data[0]);
      
      // 设置表头
      worksheet.columns = headers.map(header => ({
        header: header,
        key: header,
        width: header === '图片' ? 30 : 15
      }));

      // 添加数据行
      const row = worksheet.addRow(data[0]);
      
      // 设置行高以容纳图片
      row.height = 150; // 设置行高约150像素

      // 处理图片
      const imgPath = this.data.formData.img;
      if (imgPath) {
        try {
          // 读取图片为 base64
          const base64Data = await this.readImageFile(imgPath);
          
          // 获取图片列的索引（"图片"列）
          const imageColumnIndex = headers.indexOf('图片');
          
          if (imageColumnIndex !== -1) {
            // 添加图片到工作簿
            const imageId = workbook.addImage({
              base64: base64Data,
              extension: 'png', // 或 'jpeg'，根据实际图片类型
            });

            // 将图片插入到指定单元格
            // 行索引从 0 开始，但数据从第 2 行开始（第 1 行是表头）
            worksheet.addImage(imageId, {
              tl: { col: imageColumnIndex, row: 1 }, // 左上角位置
              ext: { width: 200, height: 200 }, // 图片尺寸
              editAs: 'oneCell'
            });

            // 清空单元格文本（因为图片已插入）
            const cell = worksheet.getCell(2, imageColumnIndex + 1);
            cell.value = '';
          }
        } catch (err) {
          console.error("插入图片失败：", err);
          // 如果插入失败，在单元格中显示提示
          const imageColumnIndex = headers.indexOf('图片');
          if (imageColumnIndex !== -1) {
            const cell = worksheet.getCell(2, imageColumnIndex + 1);
            cell.value = '图片插入失败';
          }
        }
      }

      // 生成 Excel 文件的 ArrayBuffer
      const buffer = await workbook.xlsx.writeBuffer();
      return buffer;
      
    } catch (err) {
      console.error("生成Excel失败：", err);
      throw err;
    }
  }

})
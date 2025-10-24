const XLSX = require('../../utils/xlsx.full.min.js');

Page({

  /**
   * 页面的初始数据
   */
  data: {
    testrow: [
      { 姓名: '张三', 手机: '138****0001', 反馈: '很好' },
      { 姓名: '李四', 手机: '139****0002', 反馈: '一般' }
    ]
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


  onSubmit() {
    console.log("表单已提交");
    const fs = wx.getFileSystemManager();
    const filepath = `${wx.env.USER_DATA_PATH}/test.xlsx`;
    const bookdata = this.generateXLSX(this.data.testrow);
    fs.writeFile({
      filePath: filepath,
      data: bookdata,
      encoding: 'binary',
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
  },

  // xlsx生成函数
  // 将JSON数据转换为XLSX格式
  // 并返回二进制数据
  generateXLSX(data) {
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(data);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
    const wopts = { bookType: 'xlsx', type: 'array' };
    const binaryData = XLSX.write(workbook, wopts);
    return binaryData;
  }

})
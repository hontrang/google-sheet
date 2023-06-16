function layChiSoVnIndex() {
  const ngayHienTai = moment().format("YYYY-MM-DD");
  const duLieuNgayMoiNhat = SheetUtility.layDuLieuTrongO("HOSE", "A1");
  const url = "https://wgateway-iboard.ssi.com.vn/graphql";

  const query = "query indexQuery($indexIds: [String!]!) {indexRealtimeLatestByArray(indexIds: $indexIds) {indexID indexValue allQty allValue totalQtty totalValue advances declines nochanges ceiling floor change changePercent ratioChange __typename } }";
  const variables = `{"indexIds": ["VNINDEX" ]}`;
  const object = SheetHttp.sendGraphQLRequest(url, query, variables);
  const duLieuNhanVe = object.data.indexRealtimeLatestByArray[0];

  if (duLieuNgayMoiNhat !== ngayHienTai) {
    SheetUtility.chen1HangVaoDauSheet("HOSE");
  }
  const mang_du_lieu_chinh = [[ngayHienTai, duLieuNhanVe.indexValue, duLieuNhanVe.changePercent / 100, duLieuNhanVe.totalValue * 1000000]];

  SheetUtility.ghiDuLieuVaoDayTheoTen(mang_du_lieu_chinh, "HOSE", 1, "A");
}
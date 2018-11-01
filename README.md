# zyl
public class POIExport <T> {
	
		public static <T> void downstudents(HttpServletRequest request, HttpServletResponse response,List<T> dataset,String title,String[] headers)throws IOException{
	 	
	    // 声明一个工作薄
	    HSSFWorkbook workbook = new HSSFWorkbook();
	    // 生成一个表格
	    HSSFSheet sheet = workbook.createSheet(title);
	    //新建单元格样式    
	    HSSFCellStyle cellStyle = workbook.createCellStyle();
	     //设置自动居中
	    cellStyle.setAlignment((short)2);
	    
	    // 设置表格默认列宽度为15个字节
	    sheet.setDefaultColumnWidth((short) 15);
	    
	  //这个就是合并单元格
	  //参数说明：1：开始行 2：结束行  3：开始列 4：结束列
	  //比如我要合并 第二行到第四行的    第六列到第八列     sheet.addMergedRegion(new CellRangeAddress(1,3,5,7));
	    sheet.addMergedRegion(new CellRangeAddress(0,0,0,headers.length-1));
	    HSSFRow rowTitle = sheet.createRow(0);
	    HSSFCell cellTitle = rowTitle.createCell(0);
	    cellTitle.setCellValue(title);
	    cellTitle.setCellStyle(cellStyle);
	    
	    HSSFRow row = sheet.createRow(1);
	    for (short i = 0; i < headers.length; i++) {
	        HSSFCell cell = row.createCell(i);
	        HSSFRichTextString text = new HSSFRichTextString(headers[i]);
	        cell.setCellValue(text);
	        cell.setCellStyle(cellStyle);
	    }
	    //遍历集合数据，产生数据行
	    Iterator it = dataset.iterator();
	    int index = 1;
	    int order = 0;
	    while (it.hasNext()) {
	        index++;
	        row = sheet.createRow(index);
	        T t = (T) it.next();
	        //利用反射，根据javabean属性的先后顺序，动态调用getXxx()方法得到属性值
	        Field[] fields = t.getClass().getDeclaredFields();
	        for (short i = 0; i < fields.length; i++) {
	        	  HSSFCell cell = row.createCell(i);
		            cell.setCellStyle(cellStyle);
	        	if(i==0) {
	        		  order++;
	        		  HSSFRichTextString richString = new HSSFRichTextString(""+order);
	                  HSSFFont font3 = workbook.createFont();
	                  font3.setColor(HSSFColor.BLACK.index);//定义Excel数据颜色
	                  richString.applyFont(font3);
	                  cell.setCellValue(richString);
	        	}else {
	            Field field = fields[i];
	            String fieldName = field.getName();
//	            拼接字符串获得属性的get方法
	            String getMethodName = "get"
	                    + fieldName.substring(0, 1).toUpperCase()
	                    + fieldName.substring(1);
	    
	            try {
	                Class tCls = t.getClass();
	              //getMethod第一个参数是方法名，第二个参数是该方法的参数类型，
	              //因为存在同方法名不同参数这种情况，所以只有同时指定方法名和参数类型才能唯一确定一个方法
	                Method getMethod = tCls.getMethod(getMethodName,new Class[]{});
//	                调用t类的getMethod方法，new Object[]{}是参数列表
	                Object value = getMethod.invoke(t, new Object[]{});
	                String textValue = null;

	                if (value instanceof Date)
	                {
	                    Date date = (Date) value;
	                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	                    textValue = sdf.format(date);
	                }else if(value==null) {
	                	textValue = "";
	                }
	                else
	                {
	                    //其它数据类型都当作字符串简单处理
	                    textValue = value.toString();
	                }       
//	                HSSFRichTextString 对cell中的每一段文字设置字体
	                        HSSFRichTextString richString = new HSSFRichTextString(textValue);
	                        HSSFFont font3 = workbook.createFont();
	                        font3.setColor(HSSFColor.BLACK.index);//定义Excel数据颜色
	                        richString.applyFont(font3);
	                        cell.setCellValue(richString);
	                 
	            } catch (SecurityException e) {
	                // TODO Auto-generated catch block
	                e.printStackTrace();
	            } catch (NoSuchMethodException e) {
	                // TODO Auto-generated catch block
	                e.printStackTrace();
	            } catch (IllegalArgumentException e) {
	                // TODO Auto-generated catch block
	                e.printStackTrace();
	            } catch (IllegalAccessException e) {
	                // TODO Auto-generated catch block
	                e.printStackTrace();
	            } catch (InvocationTargetException e) {
	                // TODO Auto-generated catch block
	                e.printStackTrace();
	            }
	        }
	    }
	    }
	    response.setContentType("application/octet-stream");
	    response.setHeader("Content-disposition", "attachment;filename="+java.net.URLEncoder.encode(title, "UTF-8")+".xls");//默认Excel名称
	    response.flushBuffer();
	    workbook.write(response.getOutputStream());
	   }
	}

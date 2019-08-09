% show a photo in Excel
% GitHub: github.com/Oslomayor/Pixels-in-Excel
% 2019-8-9 

[fname, fpath] = uigetfile('*.*', 'Choose a photo');
img = imread([fpath, fname]);
hight = size(img,1);
width = size(img,2);

% ActiveX to color excel cell
% Connect to Excel
Excel = actxserver('excel.application');
Excel.visible = 1;
% Get Workbook object
workbook = Excel.Workbooks.Add;
Sheets = get(workbook, 'Sheets');
Sheet1 = get(Sheets,'Item',1);
Sheet1.Activate;
% Set the color of cell "xxx" of Sheet 1 to RGB
for row = 1:hight
    for ii = 1:width
        col = num2col_recur(ii);
        cell_index = [col,num2str(row)];
        R = img(row,ii,1);
        G = img(row,ii,2);
        B = img(row,ii,3);
        % Excel µÄµ¥Ôª¸ñÑÕÉ«ÊÇBGRË³Ðò
        color_value = uint32(B)*256*256+ uint32(G)*256 + uint32(R);
        Sheet1.Range(cell_index).Interior.Color = color_value;
        disp([num2str(row),'/',num2str(hight),' ',num2str(ii),'/',num2str(width)]);
    end
end
% Save Workbook
% saveas(workbook,'result.xls');
% Close Workbook
delete(Excel);

% ShangBanMoYu

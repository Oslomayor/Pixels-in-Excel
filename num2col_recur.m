% 返回数字至 Excel 列的映射
% [col] = num2col_recur(num)
% 123...2627... -> ABC...ZAA...
% Example:
% col = num2col(28) 
% Output: col = AB

function col = num2col_recur(num)
    s= '';
    while num > 0
        s = [char(mod((num - 1),26)+65),s];
        num = fix((num - 1)/26);
    end
    col = s;
end
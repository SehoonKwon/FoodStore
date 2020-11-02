[num,txt,raw] = xlsread('원본_네이버페이.xlsx');

[row, col] = size(raw);

s1 = '배송비';


[num2,txt2,Tax_free_raw] = xlsread('Tax_free.xlsx');
[num3,txt3,Tax_raw] = xlsread('Tax.xlsx');

[row_Tax_free, col_Tax_free] = size(Tax_free_raw);
[row_Tax, col_Tax] = size(Tax_raw);

for i=3:row
    if (strcmp(raw(i,6), s1))
        for j=1 : row_Tax_free
            if(strcmp(raw(i-1, 6), Tax_free_raw(j,1)))
                temp = raw(i, 9);
                raw(i, 9) = raw(i, 10);
                raw(i, 10) = temp;
                disp(i);
                break;
            end
        end
    end
end

xlswrite('완료_네이버페이.xlsx', raw);
# KÍCH HOẠT OFFICE 2010 VÀ OFFICE 2013 #

**Tools và cmd kích hoạt giống nhau:**

## 1. AIO TOOLS V3.1.3 ##

- Download AIO Tools V3.13 [tại đây](https://1drv.ms/u/s!AkwSBX-xWiVhgReolwU8a9uuJrz7?e=AyNym8), giải nén, rồi run file Activate AIO Tools v3.1.3 by Savio.cmd bằng quyền Run Administrator

![image](https://user-images.githubusercontent.com/103977676/200758657-1ddf0219-2f35-4501-a18f-a084e5dcce15.png)

Sau khi kích hoạt xong ta gia hạn kích hoạt vĩnh viễn.

- Tạo tác vụ gia hạn kích hoạt vĩnh viến:

![image](https://user-images.githubusercontent.com/103977676/200756492-50b60776-f99b-4e12-8352-090c14850910.png)

Bấm phím O để tạo tác vụ rồi làm theo đề xuất.

- Convert Office Retail sang Office Volume: 

![image](https://user-images.githubusercontent.com/103977676/200759447-a0c844d0-04d8-4bd2-a25d-c6711b080ee4.png)

Bấm phím L

![image](https://user-images.githubusercontent.com/103977676/200759618-7f8782ce-ae9c-4f7d-bc4e-5e273bc83b13.png)

Chọn phím thích hợp rồi chuyển
 
 ## 2. Online KMS Activation Script v6.0 ##

- Download file Online KMS Activation Script v6.0.txt [tại đây](https://1drv.ms/t/s!AkwSBX-xWiVhgRK591WjSVADwexy?e=1SdXR5), rồi mở lên bấm Save As với tên kichhoat.cmd run fle này dưới quyền Run Administrator

![image](https://user-images.githubusercontent.com/103977676/200760926-e43b81b3-67e9-4949-bbe8-bc7b045a0dc6.png)

Bấm phím 2, sau đó làm theo đề xuất, kích hoạt xong tạo gia hạn kích hoạt vĩnh viễn

![image](https://user-images.githubusercontent.com/103977676/200757742-48204110-7a4d-4897-a28c-7efdedcb2fad.png)

## 3. MAS 1.5 AIO CRC32 21D20776 ##

- Download file MAS 1.5 AIO CRC32 21D20776.txt [tại đây](https://1drv.ms/t/s!AkwSBX-xWiVhgQ2uicZ7U2jSug8O?e=pd3od9), rồi mở lên bấm Save As với tên kichhoat.cmd run fle này dưới quyền Run Administrator
- Sau khi kích hoạt xong, dùng AIO TOOLS V3.1.3 hoặc Online KMS Activation Script v6.0 tạo gia hạn kích hoạt vĩnh viễn.

# KIỂM TRA TÌNH TRẠNG OFFICE THỜI GIAN KÍCH HOẠT CÒN BAO LÂU #

Dùng: AIO TOOLS V3.1.3

![image](https://user-images.githubusercontent.com/103977676/200762904-e08f2581-cdf0-46d5-b2c1-c1fef6b124cb.png)

Bấm phiếm N

![image](https://user-images.githubusercontent.com/103977676/200763673-959c1572-5c4c-42f4-9424-43fc96955838.png)

Ta xóa bỏ key Office 16lỗi có 5 ký tự cuối VMFTK

![image](https://user-images.githubusercontent.com/103977676/200764433-13ecc560-eca6-43b5-8efb-d6dc9436b35f.png)

Bấm phím M

![image](https://user-images.githubusercontent.com/103977676/200764795-ed300532-6560-49ee-b475-0d40880a78c6.png)

Bấm phím 3, điền 5 kí tự cuối vào bấm Enter

![image](https://user-images.githubusercontent.com/103977676/200765155-c5b8bc3d-135b-47bf-a94a-3cec56f638f0.png)

Kiểm tra lại ta thấy:

![image](https://user-images.githubusercontent.com/103977676/200765447-5e3c87f0-d179-4ebf-ad0b-f3f9a5fcae4c.png)

# KÍCH HOẠT BẰNG CMD CỤ THỂ #

## OFFICE 2010 ##

- Mở CMD bằng quyền Adminstrator
- ![image](https://user-images.githubusercontent.com/103977676/206964468-9b960a5a-bd66-4c96-a69a-0b58d687c7c9.png)
- Copy lệnh bên dưới và dán vào CMD

```php
if exist "%ProgramFiles%\Microsoft Office\Office14\ospp.vbs" set folder="%ProgramFiles%\Microsoft Office\Office14"
if exist "cd /d %ProgramFiles(x86)%\Microsoft Office\Office14\ospp.vbs" set folder="cd /d %ProgramFiles(x86)%\Microsoft Office\Office14"
```
- ![image](https://user-images.githubusercontent.com/103977676/206964652-58de7663-f914-421b-a590-6fadcd10847e.png)
- Đảm bảo PC của bạn được kết nối với internet, sau đó chạy lệnh sau:

```php
cscript %folder%\ospp.vbs /inpkey:VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB
cscript %folder%\ospp.vbs /inpkey:YC7DK-G2NP3-2QQC3-J6H88-GVGXT
cscript %folder%\ospp.vbs /sethst:kms8.msguides.com
cscript %folder%\ospp.vbs /setprt:1688
cscript %folder%\ospp.vbs /act
```

- ![image](https://user-images.githubusercontent.com/103977676/206964810-9c18a9e6-e29c-420f-b66f-bc385cb6f915.png)
- Nếu cmd mà hiện dòng chữ “PRODUCT ACTIVATION SUCCESSFUL” thì có nghĩa là bạn đã kích hoạt thành công Office 2010 180 ngày rồi đấy nhé!
- Bạn cần gia hạn kích hoạt chứ không hết hạn 180 ngày bạn không thể sử dụng được office!
- Trường hợp bạn bị chặn không cho sử dụng office vì hết hạn 180 ngày, đừng lo lắng, bạn tắt hết office và kích hoạt lại như trên là sử dụng được!

## OFFICE 2013 ##

- Đầu tiên chúng ta cần mở hộp thoại CMD lên:
- ![image](https://user-images.githubusercontent.com/103977676/206965374-2a256f9e-9f34-4329-879b-d59671501316.png)
- Copy đoạn lệnh sau dán vào và bấm enter:

```php
if exist "%ProgramFiles%\Microsoft Office\Office15\ospp.vbs" set folder="%ProgramFiles%\Microsoft Office\Office15" if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\ospp.vbs" set folder="%ProgramFiles(x86)%\Microsoft Office\Office15"
```

- Đảm bảo PC của bạn được kết nối với internet, sau đó chạy lệnh sau:

```php
cscript %folder%\ospp.vbs /inpkey:FN8TT-7WMH6-2D4X9-M337T-2342K
cscript %folder%\ospp.vbs /inpkey:YC7DK-G2NP3-2QQC3-J6H88-GVGXT
cscript %folder%\ospp.vbs /inpkey:KBKQT-2NMXY-JJWGP-M62JB-92CD4
cscript %folder%\ospp.vbs /inpkey:C2FG9-N6J68-H8BTJ-BW3QX-RM3B3
cscript %folder%\ospp.vbs /sethst:kms8.msguides.com
cscript %folder%\ospp.vbs /setprt:1688
cscript %folder%\ospp.vbs /act
```

- Nếu cmd mà hiện dòng chữ “PRODUCT ACTIVATION SUCCESSFUL” thì có nghĩa là bạn đã kích hoạt thành công Office 2010 180 ngày rồi đấy nhé!
- Bạn cần gia hạn kích hoạt chứ không hết hạn 180 ngày bạn không thể sử dụng được office!
- Trường hợp bạn bị chặn không cho sử dụng office vì hết hạn 180 ngày, đừng lo lắng, bạn tắt hết office và kích hoạt lại như trên là sử dụng được!

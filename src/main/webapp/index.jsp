<%@ page contentType="text/html;charset=UTF-8" %>
<%@ include file="/base/base.jsp"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
    <title>bootstrap文件上传</title>
    <script src="${ctx}/js/jquery/2.0.0/jquery.min.js"></script>
    <script src="${ctx}/js/bootstrap/3.3.6/bootstrap.min.js"></script>
    <link href="${ctx}/css/bootstrap/3.3.6/bootstrap.min.css" rel="stylesheet">
    <link type="text/css" rel="stylesheet" href="${ctx}/css/fileinput/fileinput.css"/>
    <link type="text/css" rel="stylesheet" href="${ctx}/css/fileinput/fileinput.min.css"/>
    <script type="text/javascript" src="${ctx}/js/fileinput/fileinput.js"></script>
    <script type="text/javascript" src="${ctx}/js/fileinput/fileinput.min.js"></script>
    <script type="text/javascript" src="${ctx}/js/fileinput/fileinput_locale_zh.js"></script>
    <style>
        body {
            font-family: "Arial", "Microsoft YaHei", "黑体", "宋体", sans-serif;
            background-color: #EBF5F3;
        }

        * {
            margin: 0;
        }

        html, body {
            height: 100%;
        }

        .navbar-custom {
            background-color: #56b9ab;
        }

        .navbar-brand,
        .navbar-nav li a {
            line-height: 55px;
            height: 55px;
            padding-top: 0px;
            font-family: "Arial", "Microsoft YaHei", "黑体", "宋体", sans-serif;
        }

        .navbar-default .navbar-nav > li > a {
            color: #ffffff;
        }

        .navbar-default .navbar-nav > li > a:hover {
            color: #175A94;
        }

        .page-header {
            font-family: "Arial", "Microsoft YaHei", "黑体", "宋体", sans-serif;
        }

        hr {
            border-bottom: 1px solid #bbb;
        }

        .img_border {
            border: 1px solid #bbb;
        }

        @media screen and (min-width: 800px) {
            .container {
                width: 800px;
            }
        }

        @media screen and (min-width: 800px) {
            .center_toaster {
                right: 35%;
                width: 30%;
            }
        }

        @media screen and (min-width: 400px) and (max-width: 799px) {
            .center_toaster {
                right: 25%;
                width: 50%;
            }
        }

        @media screen and (min-width: 200px ) and (max-width: 399px) {
            .center_toaster {
                right: 10%;
                width: 80%;
            }
        }

        .row a {
            text-decoration: none;
        }

        .row a:hover {
            text-decoration: none;
        }

        .wrapper {
            min-height: 100%;
            height: auto !important;
            height: 100%;
            margin: 0 auto -6em;
        }

        .push {
            height: 6em;
        }

        .footer {
            height: 4em;
        }
    </style>

    <script src="//cdn.bootcss.com/html5shiv/3.7.2/html5shiv.min.js"></script>
    <script src="//cdn.bootcss.com/respond.js/1.4.2/respond.min.js"></script>

</head>

<body>

<div class="wrapper">
    <div class="navbar navbar-default navbar-custom" role="navigation">
        <div class="navbar-header">
            <a class="navbar-brand" href="">
            </a>
        </div>
    </div>

    <div class="container kv-main">
        <div class="row ">
            <div style="padding:10px; ">
                <form enctype="multipart/form-data">
                    <input id="file-0a" class="file" name="file" type="file" multiple data-min-file-count="1">
                    <br>
                </form>
            </div>
        </div>


    </div>
    <div class="push"></div>
</div>
</body>
<footer class="footer">
    <div class=" col-lg-12 text-center">
        <hr>
        <p>北京大成股份有限公司 &copy; 2017zgzcwy.com</p>
    </div>
</footer>
    <script>
        $('#file-0a').fileinput({
            language: 'zh',
            uploadUrl: '${ctx}/uploadFile',
            allowedPreviewTypes: ['image', 'html', 'text', 'video', 'audio', 'flash']
        });
        $('#file-0a').on('fileuploaderror', function (event, data, previewId, index) {
            var form = data.form, files = data.files, extra = data.extra,
                response = data.response, reader = data.reader;
            console.log(data);
            console.log('File upload error');
        });
        $('#file-0a').on('fileerror', function (event, data) {
            console.log(data.id);
            console.log(data.index);
            console.log(data.file);
            console.log(data.reader);
            console.log(data.files);
        });
        $('#file-0a').on('fileuploaded', function (event, data, previewId, index) {
            var form = data.form, files = data.files, extra = data.extra,
                response = data.response, reader = data.reader;
            console.log('File uploaded triggered');
        });
    </script>
</html>
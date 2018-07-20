<?php

namespace app\controllers;

use Yii;
use yii\filters\AccessControl;
use yii\web\Controller;
use yii\web\Response;
use yii\filters\VerbFilter;
use app\models\LoginForm;
use app\models\ContactForm;
use PhpOffice\PhpWord\PhpWord;
use \PhpOffice\PhpWord\IOFactory as PHPWord_IOFactory;
class SiteController extends Controller
{
    /**
     * {@inheritdoc}
     */
    public function behaviors()
    {
        return [
            'access' => [
                'class' => AccessControl::className(),
                'only' => ['logout'],
                'rules' => [
                    [
                        'actions' => ['logout'],
                        'allow' => true,
                        'roles' => ['@'],
                    ],
                ],
            ],
            'verbs' => [
                'class' => VerbFilter::className(),
                'actions' => [
                    'logout' => ['post'],
                ],
            ],
        ];
    }

    /**
     * {@inheritdoc}
     */
    public function actions()
    {
        
//        echo self::phpGenerateWord();
    die;
    return [];
        return [
            'error' => [
                'class' => 'yii\web\ErrorAction',
            ],
            'captcha' => [
                'class' => 'yii\captcha\CaptchaAction',
                'fixedVerifyCode' => YII_ENV_TEST ? 'testme' : null,
            ],
        ];
    }
    
    
    /**
     * phpword
     *
     * @return string
     */
    public static function phpGenerateWord()
    {
        $PHPWord = new PHPWord();
        $str='第一个ｗｏrｄ';
        $section = $PHPWord->createSection();
        $PHPWord->addFontStyle('content', array('bold'=>false, 'size'=>11));
        $PHPWord->addFontStyle('rstyle', array('bold'=>true, 'italic'=>false, 'size'=>18,'align'=>'center'));
        $PHPWord->addFontStyle('tstyle', array('bold'=>true, 'italic'=>false, 'size'=>11));
        $PHPWord->addParagraphStyle('pstyle', array('align'=>'center', 'spaceAfter'=>100));
        $PHPWord->addFontStyle('end_style', array('italic'=>false, 'size'=>11,'align'=>'right'));
        $PHPWord->addParagraphStyle('endstyle', array('align'=>'right', 'spaceAfter'=>100));
        $title ='借款协议';
        $section->addText($title,'rstyle','pstyle');
        $section->addTextBreak(2); # 分段 2 设置段落距离
        $section->addText($str,'tstyle'); #一段内容
        // 设置表格样式
        $cellStyle = array('valign' => 'center', 'align' => 'center');
        $styleTable = array('borderSize' => 6, 'borderColor' => '000000', 'cellMargin' => 80, 'alignMent' => 'center');
        $PHPWord->addTableStyle('myTable', $styleTable);
        // 创建一个表格
        $table = $section->addTable('myTable');
        $table->addRow(300);  // 行
        $table->addCell(2300, $cellStyle)->addText('用户名'); //列
        $table->addCell(2300, $cellStyle)->addText('金额');  //列
        $table->addCell(2300, $cellStyle)->addText('期限'); //列
        $user[0]['user_id']=1;
        $user[0]['name']='zhangfei';
        $user[0]['money']=22;
        $deal['borrow_time']=22;
        $deal['time_type']=1;
        foreach ($user as $v) {
            $table->addRow(300);
            if ($v['user_id'] == 1) {
                $name = $v['name'];
            } else {
                $name = msubstr($v['name'], 0, 2);
            }
            $table->addCell(2300, $cellStyle)->addText($name);
            $table->addCell(2300, $cellStyle)->addText($v['money']);
            $btime = $deal['borrow_time'];
            if ($deal['time_type'] == 1) {
                $btime = $btime . '个月';
            } else {
                $btime = $btime . '天';
            }
            $table->addCell(2300, $cellStyle)->addText($btime);
        }
        # 设置图片样式[定位]
    
        $src = trim($_SERVER['DOCUMENT_ROOT'].'/public/img/zhang.jpg');
        $tmp_textrun = $section->createTextRun(array('indentLeft' => 2600)); 
        $imageStyle = array('width'=>110, 'height'=>110, 'position' => 'absolute', 'top' => -108, 'left' => 480, 'zIndex' => 4);
        $tmp_textrun->addImage($src, $imageStyle); 
        $path = './Upload/'.date('Ym').'/'.date('d').'/';
        if (!is_dir($path)){
            mkdir($path,0777,true);
        }
        $code= md5(uniqid());
        $fileName = strtolower($code).'.doc';
        $filePath = $path.$fileName;
        $objWriter = PHPWord_IOFactory::createWriter($PHPWord, 'Word2007');
        $objWriter->save($filePath);
        return  $filePath;
    }
    

    /**
     * Displays homepage.
     *
     * @return string
     */
    public function actionIndex()
    {
        return $this->render('index');
    }

    /**
     * Login action.
     *
     * @return Response|string
     */
    public function actionLogin()
    {
        if (!Yii::$app->user->isGuest) {
            return $this->goHome();
        }

        $model = new LoginForm();
        if ($model->load(Yii::$app->request->post()) && $model->login()) {
            return $this->goBack();
        }

        $model->password = '';
        return $this->render('login', [
            'model' => $model,
        ]);
    }

    /**
     * Logout action.
     *
     * @return Response
     */
    public function actionLogout()
    {
        Yii::$app->user->logout();

        return $this->goHome();
    }

    /**
     * Displays contact page.
     *
     * @return Response|string
     */
    public function actionContact()
    {
        $model = new ContactForm();
        if ($model->load(Yii::$app->request->post()) && $model->contact(Yii::$app->params['adminEmail'])) {
            Yii::$app->session->setFlash('contactFormSubmitted');

            return $this->refresh();
        }
        return $this->render('contact', [
            'model' => $model,
        ]);
    }

    /**
     * Displays about page.
     *
     * @return string
     */
    public function actionAbout()
    {
        return $this->render('about');
    }
}

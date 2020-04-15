using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using mshtml;

namespace deleteMyArticle_slrClub
{
    public partial class Form1 : Form
    {
        libCommon.clsUtil objUtil;

        string user_id;
        string user_pw;

        string mainURL;
        string myArticleURL;
        string delArticleURL;
        string articleURL;

        System.Collections.ArrayList articleList;

        System.Net.WebClient wc;
        WebBrowser wb;
        string nStep;//진행상황
        HtmlDocument doc;

        int progressFlag;
        int used_marketFlag;
        int myarticlePage;

        string lastItem;

        public Form1()
        {
            InitializeComponent();

            objUtil = new libCommon.clsUtil();

            user_id = objUtil.getAppCfg("user_id");
            user_pw = objUtil.getAppCfg("password");

            if(user_id.Trim().Length ==0 || user_pw.Trim().Length==0)
            {
                MessageBox.Show("로그인 정보 없음");
                Application.Exit();
            }

            mainURL = "http://www.slrclub.com";
            myArticleURL = "http://www.slrclub.com/mypage/myarticle.php?";
            delArticleURL = "http://www.slrclub.com/bbs/delete.php?";
            articleURL = "http://www.slrclub.com/bbs/vx2.php?";

            articleList = new System.Collections.ArrayList();

            wc = new System.Net.WebClient();
            wb = new WebBrowser();
            wb.ScriptErrorsSuppressed = true;
            wb.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(wb_DocumentCompleted);
            wb.DocumentText = "";
            doc = wb.Document.OpenNew(true);

            nStep = "start";

            progressFlag = 0;
            used_marketFlag = 0;
            myarticlePage = 1;

            lastItem = "";

            this.label1.Text = "";
        }

        private void setState()
        {
            string state = string.Format("페이지:{0}\r\n게시물 번호:{1}\r\n장터게시물 갯수:{2}\r\n상태:{3}", myarticlePage.ToString(), progressFlag.ToString(), used_marketFlag.ToString(), nStep);
            libMyUtil.clsThread.SetLabel(this.label1, state);            
        }

        void wb_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            setState();

            if (nStep.Equals("start"))
                login();
            else if (nStep.Equals("trylogin"))
            {
                if (login_state())
                    gotoMyArticlePage();
                else
                    MessageBox.Show("로그인 실패");
            }
            else if (nStep.Equals("gotomyarticle"))
            {
                parseMyArticlePage();
                if (isDupeLastItem())
                {
                    myarticlePage++;
                    gotoMyArticlePage();
                }
                else
                    deletelist();
            }
            else if (nStep.Equals("delete"))
            {
                progressFlag++;
                confirmDelete();
            }
            else if (nStep.Equals("confirmdelete"))
            {
                //삭제 확인 버튼 누른 후 여기로 이동
                deletelist();
            }
        }

        private bool isDupeLastItem()
        {
            if (articleList.Count > 0)
            {
                if (((Article)articleList[articleList.Count - 1]).articleNo.Equals(lastItem))
                    return true;
                else
                    return false;
            }
            else
                return false;
        }

        private void confirmDelete()
        {
            int flag = 0;

            /*
            if (doc.GetElementById("slrct") != null && doc.GetElementById("slrct").InnerText.IndexOf("삭제하시겠습니까?") > -1)
            {
                mshtml.HTMLDocument hdom = doc.DomDocument as mshtml.HTMLDocument;

                foreach (mshtml.IHTMLElement hel in (mshtml.IHTMLElementCollection)hdom.body.all)
                {
                    if (hel.tagName.ToLower().Equals("input") && hel.getAttribute("type").ToString().ToLower().Equals("image") && hel.getAttribute("accesskey").ToString().ToLower().Equals("s") && hel.getAttribute("alt").ToString().Equals("확인"))
                    {
                        nStep = "confirmdelete";
                        flag++;
                        hel.click();
                        break;
                    }
                }

                if (flag == 0)
                    MessageBox.Show("확인버튼 못찾음");
            }
            else
                deletelist();
             */

            mshtml.HTMLDocument hdom = doc.DomDocument as mshtml.HTMLDocument;

            foreach (mshtml.IHTMLElement hel in (mshtml.IHTMLElementCollection)hdom.body.all)
            {
                if (hel.tagName.ToLower().Equals("input") && hel.getAttribute("type").ToString().ToLower().Equals("image") && hel.getAttribute("alt").ToString().Equals("확인"))
                {
                    nStep = "confirmdelete";
                    flag++;
                    hel.click();
                    break;
                }
            }

            if (flag == 0)
            {
                //MessageBox.Show("확인버튼 못찾음");
                deletelist();
            }
        }

        private void deletelist()
        {
            nStep = "delete";

            string bbs;
            string no;
            bool isMarket = true;

            if (progressFlag < articleList.Count)
            {
                while (isMarket)
                {
                    bbs = ((Article)articleList[progressFlag]).articleBBS;
                    no = ((Article)articleList[progressFlag]).articleNo;

                    if (!bbs.Equals("used_market"))
                    {
                        string targetpage = string.Format("{0}id={1}&no={2}", articleURL, bbs, no);
                        string deletepage = string.Format("{0}id={1}&no={2}", delArticleURL, bbs, no);

                        wb.Navigate(deletepage, null, null, string.Format("Referer:{0}", targetpage));
                        isMarket = false;
                    }
                    else
                    {
                        isMarket = true;
                        used_marketFlag++;
                        progressFlag++;
                    }

                    lastItem = no;

                    if (progressFlag == articleList.Count)
                        break;
                }
            }
            
            if(progressFlag >= articleList.Count)
                gotoMyArticlePage();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            libMyUtil.clsThread.SetLabel(this.label1, "준비중...");
            wb.Navigate(mainURL);
        }

        private void parseMyArticlePage()
        {
            articleList.RemoveRange(0, articleList.Count);
            progressFlag = 0;
            used_marketFlag = 0;

            string href;
            string bbs;
            string no;
            int i;

            for (i = 0; i < doc.GetElementsByTagName("tr").Count; i++)
            {
                if (doc.GetElementsByTagName("tr")[i].InnerText != null)
                {
                    if (doc.GetElementsByTagName("tr")[i].GetElementsByTagName("td").Count == 3)
                    {
                        if (doc.GetElementsByTagName("tr")[i].GetElementsByTagName("td")[0].GetElementsByTagName("a").Count == 1)
                        {
                            if (doc.GetElementsByTagName("tr")[i].GetElementsByTagName("td")[2].InnerText != null)
                            {
                                if (doc.GetElementsByTagName("tr")[i].GetElementsByTagName("td")[2].InnerText.Length == 19)
                                {
                                    HtmlElement htmlElem = doc.GetElementsByTagName("tr")[i].GetElementsByTagName("td")[0].GetElementsByTagName("a")[0];
                                    href = htmlElem.GetAttribute("href");

                                    string[] tmp = objUtil.Split(href.Substring(href.IndexOf("?") + 1), "&");

                                    bbs = objUtil.Split(tmp[0], "=")[1];
                                    no = objUtil.Split(tmp[1], "=")[1];

                                    articleList.Add(new Article(no, bbs));
                                }
                            }
                        }
                    }
                }
            }
        }

        //로그인 과정
        private void login()
        {
            if (doc.GetElementById("login-input") != null)
            {
                nStep = "trylogin";

                doc.GetElementById("id-pw").GetElementsByTagName("input").GetElementsByName("user_id")[0].SetAttribute("value", user_id);
                doc.GetElementById("id-pw").GetElementsByTagName("input").GetElementsByName("password")[0].SetAttribute("value", user_pw);

                mshtml.IHTMLElement hel = doc.GetElementById("login-process").GetElementsByTagName("input")[0].DomElement as mshtml.IHTMLElement;
                hel.click();
            }
        }

        private void gotoMyArticlePage()
        {
            nStep = "gotomyarticle";
            wb.Navigate(string.Format("{0}page={1}", myArticleURL, myarticlePage));
        }

        private bool login_state()
        {
            if (doc.GetElementById("login-input") == null)
                return true;
            else
                return false;
        }

        public class Article
        {
            public string articleNo;
            public string articleBBS;

            public Article(string no, string bbs)
            {
                articleNo = no;
                articleBBS = bbs;
            }
        }        
    }
}

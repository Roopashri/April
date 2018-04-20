//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.34209
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

// 
// This source code was auto-generated by Microsoft.VSDesigner, Version 4.0.30319.34209.
// 
#pragma warning disable 1591

namespace OCRSWinServiceOfflineOrder.WSResearchBPM {
    using System;
    using System.Web.Services;
    using System.Diagnostics;
    using System.Web.Services.Protocols;
    using System.Xml.Serialization;
    using System.ComponentModel;
    
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.0.30319.34209")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Web.Services.WebServiceBindingAttribute(Name="RBPMWebServiceSoapBinding", Namespace="urn:RBPMWebService")]
    public partial class RBPMWebServiceInterfaceService : System.Web.Services.Protocols.SoapHttpClientProtocol {
        
        private System.Threading.SendOrPostCallback createOnlineCaseOperationCompleted;
        
        private System.Threading.SendOrPostCallback updateOnlineCaseOperationCompleted;
        
        private System.Threading.SendOrPostCallback downloadOnlineReportOperationCompleted;
        
        private System.Threading.SendOrPostCallback cancelOnlineOrderOperationCompleted;
        
        private bool useDefaultCredentialsSetExplicitly;
        
        /// <remarks/>
        public RBPMWebServiceInterfaceService() {
            this.Url = "http://202.56.56.246:8181/Research-BPM/services/RBPMWebService";
            if ((this.IsLocalFileSystemWebService(this.Url) == true)) {
                this.UseDefaultCredentials = true;
                this.useDefaultCredentialsSetExplicitly = false;
            }
            else {
                this.useDefaultCredentialsSetExplicitly = true;
            }
        }
        
        public new string Url {
            get {
                return base.Url;
            }
            set {
                if ((((this.IsLocalFileSystemWebService(base.Url) == true) 
                            && (this.useDefaultCredentialsSetExplicitly == false)) 
                            && (this.IsLocalFileSystemWebService(value) == false))) {
                    base.UseDefaultCredentials = false;
                }
                base.Url = value;
            }
        }
        
        public new bool UseDefaultCredentials {
            get {
                return base.UseDefaultCredentials;
            }
            set {
                base.UseDefaultCredentials = value;
                this.useDefaultCredentialsSetExplicitly = true;
            }
        }
        
        /// <remarks/>
        public event createOnlineCaseCompletedEventHandler createOnlineCaseCompleted;
        
        /// <remarks/>
        public event updateOnlineCaseCompletedEventHandler updateOnlineCaseCompleted;
        
        /// <remarks/>
        public event downloadOnlineReportCompletedEventHandler downloadOnlineReportCompleted;
        
        /// <remarks/>
        public event cancelOnlineOrderCompletedEventHandler cancelOnlineOrderCompleted;
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapRpcMethodAttribute("", RequestNamespace="urn:RBPMWebService", ResponseNamespace="urn:RBPMWebService")]
        [return: System.Xml.Serialization.SoapElementAttribute("createOnlineCaseReturn")]
        public string createOnlineCase(string in0) {
            object[] results = this.Invoke("createOnlineCase", new object[] {
                        in0});
            return ((string)(results[0]));
        }
        
        /// <remarks/>
        public void createOnlineCaseAsync(string in0) {
            this.createOnlineCaseAsync(in0, null);
        }
        
        /// <remarks/>
        public void createOnlineCaseAsync(string in0, object userState) {
            if ((this.createOnlineCaseOperationCompleted == null)) {
                this.createOnlineCaseOperationCompleted = new System.Threading.SendOrPostCallback(this.OncreateOnlineCaseOperationCompleted);
            }
            this.InvokeAsync("createOnlineCase", new object[] {
                        in0}, this.createOnlineCaseOperationCompleted, userState);
        }
        
        private void OncreateOnlineCaseOperationCompleted(object arg) {
            if ((this.createOnlineCaseCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.createOnlineCaseCompleted(this, new createOnlineCaseCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapRpcMethodAttribute("", RequestNamespace="urn:RBPMWebService", ResponseNamespace="urn:RBPMWebService")]
        [return: System.Xml.Serialization.SoapElementAttribute("updateOnlineCaseReturn")]
        public string updateOnlineCase(string in0, string in1) {
            object[] results = this.Invoke("updateOnlineCase", new object[] {
                        in0,
                        in1});
            return ((string)(results[0]));
        }
        
        /// <remarks/>
        public void updateOnlineCaseAsync(string in0, string in1) {
            this.updateOnlineCaseAsync(in0, in1, null);
        }
        
        /// <remarks/>
        public void updateOnlineCaseAsync(string in0, string in1, object userState) {
            if ((this.updateOnlineCaseOperationCompleted == null)) {
                this.updateOnlineCaseOperationCompleted = new System.Threading.SendOrPostCallback(this.OnupdateOnlineCaseOperationCompleted);
            }
            this.InvokeAsync("updateOnlineCase", new object[] {
                        in0,
                        in1}, this.updateOnlineCaseOperationCompleted, userState);
        }
        
        private void OnupdateOnlineCaseOperationCompleted(object arg) {
            if ((this.updateOnlineCaseCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.updateOnlineCaseCompleted(this, new updateOnlineCaseCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapRpcMethodAttribute("", RequestNamespace="urn:RBPMWebService", ResponseNamespace="urn:RBPMWebService")]
        [return: System.Xml.Serialization.SoapElementAttribute("downloadOnlineReportReturn", DataType="base64Binary")]
        public byte[] downloadOnlineReport(string in0, string in1, float in2) {
            object[] results = this.Invoke("downloadOnlineReport", new object[] {
                        in0,
                        in1,
                        in2});
            return ((byte[])(results[0]));
        }
        
        /// <remarks/>
        public void downloadOnlineReportAsync(string in0, string in1, float in2) {
            this.downloadOnlineReportAsync(in0, in1, in2, null);
        }
        
        /// <remarks/>
        public void downloadOnlineReportAsync(string in0, string in1, float in2, object userState) {
            if ((this.downloadOnlineReportOperationCompleted == null)) {
                this.downloadOnlineReportOperationCompleted = new System.Threading.SendOrPostCallback(this.OndownloadOnlineReportOperationCompleted);
            }
            this.InvokeAsync("downloadOnlineReport", new object[] {
                        in0,
                        in1,
                        in2}, this.downloadOnlineReportOperationCompleted, userState);
        }
        
        private void OndownloadOnlineReportOperationCompleted(object arg) {
            if ((this.downloadOnlineReportCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.downloadOnlineReportCompleted(this, new downloadOnlineReportCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapRpcMethodAttribute("", RequestNamespace="urn:RBPMWebService", ResponseNamespace="urn:RBPMWebService")]
        [return: System.Xml.Serialization.SoapElementAttribute("cancelOnlineOrderReturn")]
        public string cancelOnlineOrder(string in0) {
            object[] results = this.Invoke("cancelOnlineOrder", new object[] {
                        in0});
            return ((string)(results[0]));
        }
        
        /// <remarks/>
        public void cancelOnlineOrderAsync(string in0) {
            this.cancelOnlineOrderAsync(in0, null);
        }
        
        /// <remarks/>
        public void cancelOnlineOrderAsync(string in0, object userState) {
            if ((this.cancelOnlineOrderOperationCompleted == null)) {
                this.cancelOnlineOrderOperationCompleted = new System.Threading.SendOrPostCallback(this.OncancelOnlineOrderOperationCompleted);
            }
            this.InvokeAsync("cancelOnlineOrder", new object[] {
                        in0}, this.cancelOnlineOrderOperationCompleted, userState);
        }
        
        private void OncancelOnlineOrderOperationCompleted(object arg) {
            if ((this.cancelOnlineOrderCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.cancelOnlineOrderCompleted(this, new cancelOnlineOrderCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        public new void CancelAsync(object userState) {
            base.CancelAsync(userState);
        }
        
        private bool IsLocalFileSystemWebService(string url) {
            if (((url == null) 
                        || (url == string.Empty))) {
                return false;
            }
            System.Uri wsUri = new System.Uri(url);
            if (((wsUri.Port >= 1024) 
                        && (string.Compare(wsUri.Host, "localHost", System.StringComparison.OrdinalIgnoreCase) == 0))) {
                return true;
            }
            return false;
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.0.30319.34209")]
    public delegate void createOnlineCaseCompletedEventHandler(object sender, createOnlineCaseCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.0.30319.34209")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class createOnlineCaseCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal createOnlineCaseCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public string Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((string)(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.0.30319.34209")]
    public delegate void updateOnlineCaseCompletedEventHandler(object sender, updateOnlineCaseCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.0.30319.34209")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class updateOnlineCaseCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal updateOnlineCaseCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public string Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((string)(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.0.30319.34209")]
    public delegate void downloadOnlineReportCompletedEventHandler(object sender, downloadOnlineReportCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.0.30319.34209")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class downloadOnlineReportCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal downloadOnlineReportCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public byte[] Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((byte[])(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.0.30319.34209")]
    public delegate void cancelOnlineOrderCompletedEventHandler(object sender, cancelOnlineOrderCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.0.30319.34209")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class cancelOnlineOrderCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal cancelOnlineOrderCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public string Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((string)(this.results[0]));
            }
        }
    }
}

#pragma warning restore 1591
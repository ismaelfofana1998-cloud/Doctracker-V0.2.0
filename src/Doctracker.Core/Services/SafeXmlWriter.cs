using System;
using System.Text;
using System.Xml;

namespace Doctracker.Core.Services
{
    // PDF text can contain control characters that XmlSerializer writes as &#x1;,
    // but a conforming XML 1.0 reader then rejects. Sanitize text, not markup.
    internal sealed class SafeXmlWriter : XmlWriter
    {
        private readonly XmlWriter writer;
        public SafeXmlWriter(XmlWriter writer) { this.writer=writer; }
        internal static string Clean(string value)
        {
            if(string.IsNullOrEmpty(value))return value;
            var result=new StringBuilder(value.Length);
            for(var i=0;i<value.Length;i++)
            {
                var c=value[i];
                if(char.IsHighSurrogate(c) && i+1<value.Length && char.IsLowSurrogate(value[i+1])) {result.Append(c).Append(value[++i]);continue;}
                if(XmlConvert.IsXmlChar(c))result.Append(c);
            }
            return result.ToString();
        }
        public override void WriteString(string text)=>writer.WriteString(Clean(text));
        public override void WriteChars(char[] buffer,int index,int count)=>WriteString(new string(buffer,index,count));
        public override void WriteCData(string text)=>writer.WriteCData(Clean(text));
        public override void WriteCharEntity(char ch){if(XmlConvert.IsXmlChar(ch))writer.WriteCharEntity(ch);}
        public override void WriteRaw(string data)=>writer.WriteRaw(data);
        public override void WriteRaw(char[] buffer,int index,int count)=>writer.WriteRaw(buffer,index,count);
        public override void WriteStartDocument()=>writer.WriteStartDocument();
        public override void WriteStartDocument(bool standalone)=>writer.WriteStartDocument(standalone);
        public override void WriteEndDocument()=>writer.WriteEndDocument();
        public override void WriteDocType(string name,string pubid,string sysid,string subset)=>writer.WriteDocType(name,pubid,sysid,subset);
        public override void WriteStartElement(string prefix,string localName,string ns)=>writer.WriteStartElement(prefix,localName,ns);
        public override void WriteEndElement()=>writer.WriteEndElement();
        public override void WriteFullEndElement()=>writer.WriteFullEndElement();
        public override void WriteStartAttribute(string prefix,string localName,string ns)=>writer.WriteStartAttribute(prefix,localName,ns);
        public override void WriteEndAttribute()=>writer.WriteEndAttribute();
        public override void WriteComment(string text)=>writer.WriteComment(Clean(text));
        public override void WriteProcessingInstruction(string name,string text)=>writer.WriteProcessingInstruction(name,text);
        public override void WriteEntityRef(string name)=>writer.WriteEntityRef(name);
        public override void WriteWhitespace(string ws)=>writer.WriteWhitespace(ws);
        public override void WriteSurrogateCharEntity(char lowChar,char highChar)=>writer.WriteSurrogateCharEntity(lowChar,highChar);
        public override void WriteBase64(byte[] buffer,int index,int count)=>writer.WriteBase64(buffer,index,count);
        public override WriteState WriteState=>writer.WriteState;
        public override string LookupPrefix(string ns)=>writer.LookupPrefix(ns);
        public override void Flush()=>writer.Flush();
        public override void Close()=>writer.Close();
    }
}

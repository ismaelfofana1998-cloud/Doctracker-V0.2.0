using System;
namespace Doctracker.Core.Services
{
    public static class OcrConcurrencyPolicy
    {
        public static int Maximum(int logicalProcessors)=>Math.Max(1,Math.Min(16,logicalProcessors));
        public static int Recommended(int logicalProcessors,double memoryGb)
        {
            var cpu=Math.Max(1,logicalProcessors/2);
            var memory=memoryGb<=0?2:memoryGb<6?1:memoryGb<12?2:4;
            return Math.Min(Maximum(logicalProcessors),Math.Min(cpu,memory));
        }
        public static int Clamp(int requested,int logicalProcessors)=>Math.Max(1,Math.Min(Maximum(logicalProcessors),requested));
    }
}

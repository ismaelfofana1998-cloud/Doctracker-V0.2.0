using Doctracker.Core.Services;
using Xunit;

namespace Doctracker.Core.Tests
{
    public sealed class OcrConcurrencyTests
    {
        [Theory]
        [InlineData(12,16,4,12)]
        [InlineData(12,8,2,12)]
        [InlineData(8,4,1,8)]
        [InlineData(2,16,1,2)]
        [InlineData(32,64,4,16)]
        [InlineData(12,0,2,12)]
        [InlineData(0,0,1,1)]
        public void Advice_balances_memory_cpu_and_a_conservative_start(int logical,double memory,int recommended,int maximum)
        {
            Assert.Equal(recommended,OcrConcurrencyPolicy.Recommended(logical,memory));
            Assert.Equal(maximum,OcrConcurrencyPolicy.Maximum(logical));
        }
        [Theory]
        [InlineData(-1,12,1)]
        [InlineData(0,12,1)]
        [InlineData(12,12,12)]
        [InlineData(50,12,12)]
        [InlineData(50,32,16)]
        public void Saved_settings_are_bounded_on_the_current_computer(int requested,int logical,int expected)
        {Assert.Equal(expected,OcrConcurrencyPolicy.Clamp(requested,logical));}
    }
}

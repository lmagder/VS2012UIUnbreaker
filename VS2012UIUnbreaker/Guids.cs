// Guids.cs
// MUST match guids.h
using System;

namespace Hacks.VS2012UIUnbreaker
{
    static class GuidList
    {
        public const string guidVS2012UIUnbreakerPkgString = "ec8cfd06-120d-448d-a17c-bf91d6deaf30";
        public const string guidVS2012UIUnbreakerCmdSetString = "d4428329-efdc-4cf4-9789-34026a1351d9";

        public static readonly Guid guidVS2012UIUnbreakerCmdSet = new Guid(guidVS2012UIUnbreakerCmdSetString);
    };
}
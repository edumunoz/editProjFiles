// Guids.cs
// MUST match guids.h
using System;

namespace rdomunozcom.EditProj
{
    static class GuidList
    {
        public const string guidEditProjPkgString = "67374493-7f41-4665-bb0f-9ce9ede3fe7b";
        public const string guidEditProjCmdSetString = "d2f70dae-9a2d-47e1-a470-7354a552821c";

        public static readonly Guid guidEditProjCmdSet = new Guid(guidEditProjCmdSetString);
    };
}

msbuild -restore -t:Rebuild -p:Configuration="Release" -p:Platform=x64  ^
    SampleView.sln

msbuild -restore -t:Rebuild -p:Configuration="Release" -p:Platform=x86  ^
    SampleView.sln

import unitypack

with open('D:\\misc\\Hearthstone\\Data\Win\\cardxml0.unity3d') as f:
    bundle = unitypack.load(f)

    #for asset in bundle.assets:
    #    print("%s: %s:: %i objects" % (bundle, asset, len(asset.objects)))

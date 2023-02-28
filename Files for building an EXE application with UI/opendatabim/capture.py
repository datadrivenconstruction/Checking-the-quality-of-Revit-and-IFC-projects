import sys
import collada
import numpy as np
from OCC.Core.MeshDS import MeshDS_DataSource
from OCC.Core.MeshVS import *
from OCC.Core.Quantity import *
from OCC.Core.Graphic3d import *
from OCC.Core.V3d import *
from OCC.Core.gp import *
from OCC.Display.OCCViewer import Viewer3d
from math import cos, sin, radians



background_color =  (255,255,255)
size =  (1000,1000)

class ImgCapture():

    def __init__(self):

        self.display = Viewer3d()
        self.display.Create(create_default_lights=False, display_glinfo=False)
        self.display.SetModeShaded()
        self.display.SetSize(size[0], size[1])

        self._set_settings()

    def _set_settings(self):

        # CAD Assistant Default Settings

        # Directional Light smooth angle, degrees
        smooth_angle = radians(6)
        ### spherical coordinate system ###
        # Directional Light θ, degrees
        d1 = radians(-77)
        # Directional Light φ, degrees
        d2 = radians(-153)

        # Directional Light Intensity
        dl_Intensity = 0.8
        # Ambient Light Intensity
        al_Intensity = 0.7

        R = 1
        X = R * sin(d1) * cos(d2)
        Y = R * sin(d1) * sin(d2)
        Z = R * cos(d1)

        DirectionalLight = V3d_DirectionalLight(
            gp_Dir(X, Y, Z),
            Quantity_Color(Quantity_NOC_WHITE),
            True
        )

        DirectionalLight.SetIntensity(dl_Intensity)
        DirectionalLight.SetSmoothAngle(smooth_angle)

        AmbientLight = V3d_AmbientLight()
        AmbientLight.SetIntensity(al_Intensity)


        bc = background_color
        bc = Quantity_Color(bc[0] / 255, bc[1] / 255, bc[2] / 255, Quantity_TOC_RGB)

        self.display.Viewer.AddLight(DirectionalLight)
        self.display.Viewer.AddLight(AmbientLight)
        self.display.View.SetLightOn()
        self.display.View_Iso()
        self.display.View.SetBackgroundColor(bc)

    def load(self, dae):
        c = collada.Collada(dae, ignore=[collada.common.DaeUnsupportedError, collada.common.DaeBrokenRefError])

        material_map = {}
        for m in c.materials:
            material_map[m.id] = self._parse_material(m)

        self.meshes = {}
        for node in c.scene.nodes:
            node_id = node.id
            self._parse_node(node_id=node_id,
                         node=node,
                         material_map=material_map,
                         meshes=self.meshes,
                         )


    def _parse_node(self, node_id, node, material_map, meshes):
        if isinstance(node, collada.scene.GeometryNode):
            geometry = node.geometry
            material = None
            for mn in node.materials:
                m = mn.target
                if m.id in material_map:
                    material = material_map[m.id]

            for i, primitive in enumerate(geometry.primitives):
                if isinstance(primitive, collada.polylist.Polylist):
                    primitive = primitive.triangleset()
                if isinstance(primitive, collada.triangleset.TriangleSet):
                    vertex = primitive.vertex
                    vertex_index = primitive.vertex_index
                    if vertex_index is not None:
                        vertices = vertex[vertex_index].reshape(
                            len(vertex_index) * 3, 3)

                        faces = np.arange(
                            vertices.shape[0], dtype=np.int32).reshape(
                            vertices.shape[0] // 3, 3)

                        MeshDS = MeshDS_DataSource(vertices, faces)
                        MeshVS = MeshVS_Mesh()
                        MeshVS.SetDataSource(MeshDS)

                        prs_builder = MeshVS_MeshPrsBuilder(MeshVS)

                        drawer = prs_builder.GetDrawer()
                        drawer.SetBoolean(MeshVS_DA_ShowEdges, False)
                        drawer.SetBoolean(MeshVS_DA_DisplayNodes, False)
                        drawer.SetMaterial(MeshVS_DA_FrontMaterial, material)

                        prs_builder.SetDrawer(drawer)
                        MeshVS.AddBuilder(prs_builder, True)
                        MeshVS.SetDisplayMode(MeshVS_DMF_Shading)

                        meshes[node_id] = MeshVS


        elif isinstance(node, collada.scene.Node):
            if node.children is not None:
                for child in node.children:
                    self._parse_node(
                        node_id=node.id,
                        node=child,
                        material_map=material_map,
                        meshes=meshes
                    )

    def _parse_material(self, m):

        effect = m.effect

        PBRMaterial = Graphic3d_PBRMaterial()
        material = Graphic3d_MaterialAspect(Graphic3d_NOM_SILVER)
        material.SetPBRMaterial(PBRMaterial)


        dc = np.ones(3)
        if effect.diffuse:
            dc = effect.diffuse
        diffuseColor = Quantity_Color(dc[0], dc[1], dc[2], Quantity_TOC_RGB)
        material.SetDiffuseColor(diffuseColor)

        ac = np.ones(3)
        if effect.ambient:
            ac = effect.ambient
        ambientColor = Quantity_Color(ac[0], ac[1], ac[2], Quantity_TOC_RGB)
        material.SetAmbientColor(ambientColor)

        shininess = 1.0
        if effect.shininess:
            shininess = effect.shininess
        material.SetShininess(1 - shininess)

        transparency = 1.0
        if effect.transparency:
            transparency = effect.transparency
        material.SetTransparency(1 - transparency)

        if effect.specular:
            try:
                sc = effect.specular
                specularColor = Quantity_Color(sc[0], sc[1], sc[2], Quantity_TOC_RGB)
                material.SetSpecularColor(specularColor)
            except: pass

        if effect.emission:
            try:
                ec = effect.emission
                emissiveColor = Quantity_Color(ec[0], ec[1], ec[2], Quantity_TOC_RGB)
                material.SetEmissiveColor(emissiveColor)
            except: pass


        if effect.index_of_refraction:
            try:material.SetRefractionIndex(1 - effect.index_of_refraction)
            except:pass

        return material

    def shoot(self, path, nodes=[]):
        imgHave = False

        if len(nodes)==0:
            imgHave = True
            for node in self.meshes.values():
                self.display.Context.Display(node, False)

        for node in nodes:
            # node = str(node)
            if node in self.meshes:
                imgHave = True
                self.display.Context.Display(self.meshes[node], False)

        if imgHave:
            self.display.FitAll()
            # self.start_display()
            self.display.View_Iso()
            self.display.ExportToImage(path)
            self.display.EraseAll()
            return True

        return False
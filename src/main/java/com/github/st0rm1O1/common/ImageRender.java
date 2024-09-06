package com.github.st0rm1O1.common;

import com.github.st0rm1O1.resource.Resource;
import lombok.AllArgsConstructor;

import javax.swing.*;
import javax.swing.plaf.basic.BasicPanelUI;
import java.awt.*;


@AllArgsConstructor
public class ImageRender extends BasicPanelUI {
	private final String path;

	public void paint(Graphics g, JComponent c) {
		Graphics2D g2D = (Graphics2D) g;
        g2D.setRenderingHint(
                RenderingHints.KEY_ANTIALIASING,
                RenderingHints.VALUE_ANTIALIAS_ON);
        g2D.setRenderingHint(
                RenderingHints.KEY_INTERPOLATION,
                RenderingHints.VALUE_INTERPOLATION_BILINEAR);
		g2D.drawImage(new ImageIcon(Resource.loadImage(path)).getImage(), 0, 0, null);
		g2D.dispose();
    }

    public Dimension getPreferredSize(JComponent c) {
        return super.getPreferredSize(c);
    }
	
}

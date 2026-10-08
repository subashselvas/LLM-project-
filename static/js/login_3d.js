/**
 * EduCore 3D Login Background
 * Scene with small orbiting planets and star particles (No core Jupiter planet).
 * Usage: call initLoginScene(canvasId, accentColorHex)
 */
function initLoginScene(canvasId, accentColor, showPlanets = false) {
    if (window.innerWidth <= 768) return;

    const canvas = document.getElementById(canvasId);
    if (!canvas) return;

    // Full-screen (fixed) canvas uses window dims; embedded uses parent dims
    const isFixed = getComputedStyle(canvas).position === 'fixed';
    const W = isFixed ? window.innerWidth : (canvas.offsetWidth || canvas.parentElement.offsetWidth);
    const H = isFixed ? window.innerHeight : (canvas.offsetHeight || canvas.parentElement.offsetHeight);

    const renderer = new THREE.WebGLRenderer({ canvas: canvas, alpha: true, antialias: true });
    renderer.setSize(W, H);
    renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));

    const scene = new THREE.Scene();

    const camera = new THREE.PerspectiveCamera(60, W / H, 0.1, 1000);
    camera.position.set(0, 0, 22);

    // ----- Lighting -----
    const ambient = new THREE.AmbientLight(0xffffff, 0.3);
    scene.add(ambient);

    const pointLight = new THREE.PointLight(accentColor, 4, 50);
    pointLight.position.set(0, 0, 10);
    scene.add(pointLight);

    const rimLight = new THREE.PointLight(0xffffff, 1.5, 40);
    rimLight.position.set(10, 10, -5);
    scene.add(rimLight);

    // ----- Orbiting Small Planets (Nodes) -----
    const nodeData = [];
    if (showPlanets) {
        const nodeColors = [accentColor, 0xffffff, 0xaaccff];
        const nodeCount = 24;

        for (let i = 0; i < nodeCount; i++) {
            const radius = 5 + Math.random() * 5;
            const theta = Math.random() * Math.PI * 2;
            const phi = Math.acos(2 * Math.random() - 1);

            const size = 0.08 + Math.random() * 0.18;
            const geo = Math.random() > 0.5
                ? new THREE.SphereGeometry(size, 8, 8)
                : new THREE.OctahedronGeometry(size);
            const mat = new THREE.MeshStandardMaterial({
                color: nodeColors[Math.floor(Math.random() * nodeColors.length)],
                emissive: nodeColors[Math.floor(Math.random() * nodeColors.length)],
                emissiveIntensity: 0.6,
                roughness: 0.1,
                metalness: 0.9
            });
            const node = new THREE.Mesh(geo, mat);

            node.position.set(
                radius * Math.sin(phi) * Math.cos(theta),
                radius * Math.sin(phi) * Math.sin(theta),
                radius * Math.cos(phi)
            );

            const speed = (0.001 + Math.random() * 0.003) * (Math.random() > 0.5 ? 1 : -1);
            const axis = new THREE.Vector3(
                Math.random() - 0.5,
                Math.random() - 0.5,
                Math.random() - 0.5
            ).normalize();

            scene.add(node);
            nodeData.push({ mesh: node, speed, axis, initPos: node.position.clone(), radius });
        }
    }

    // ----- Small Stars (main field) -----
    const starCount = 350;
    const starPositions = new Float32Array(starCount * 3);
    for (let i = 0; i < starCount; i++) {
        starPositions[i * 3]     = (Math.random() - 0.5) * 60;
        starPositions[i * 3 + 1] = (Math.random() - 0.5) * 60;
        starPositions[i * 3 + 2] = (Math.random() - 0.5) * 30;
    }
    const starGeo = new THREE.BufferGeometry();
    starGeo.setAttribute('position', new THREE.BufferAttribute(starPositions, 3));
    const starMat = new THREE.PointsMaterial({
        color: 0xffffff,
        size: 0.10,
        transparent: true,
        opacity: 0.7
    });
    const stars = new THREE.Points(starGeo, starMat);
    scene.add(stars);

    // ----- Accent-coloured micro stars -----
    const accentStarCount = 80;
    const accentPos = new Float32Array(accentStarCount * 3);
    for (let i = 0; i < accentStarCount; i++) {
        accentPos[i * 3]     = (Math.random() - 0.5) * 60;
        accentPos[i * 3 + 1] = (Math.random() - 0.5) * 60;
        accentPos[i * 3 + 2] = (Math.random() - 0.5) * 30;
    }
    const accentGeo = new THREE.BufferGeometry();
    accentGeo.setAttribute('position', new THREE.BufferAttribute(accentPos, 3));
    const accentMat = new THREE.PointsMaterial({
        color: accentColor,
        size: 0.08,
        transparent: true,
        opacity: 0.55
    });
    const accentStars = new THREE.Points(accentGeo, accentMat);
    scene.add(accentStars);

    // ----- Mouse Tracking -----
    let mouseX = 0, mouseY = 0;
    const el = canvas.parentElement;
    el.addEventListener('mousemove', e => {
        const rect = el.getBoundingClientRect();
        mouseX = ((e.clientX - rect.left) / rect.width - 0.5) * 2;
        mouseY = -((e.clientY - rect.top) / rect.height - 0.5) * 2;
    });

    // ----- Animate -----
    let t = 0;
    function animate() {
        requestAnimationFrame(animate);
        t += 0.008;

        // Orbit nodes (small planets)
        if (showPlanets) {
            nodeData.forEach(n => {
                n.mesh.position.applyAxisAngle(n.axis, n.speed);
                n.mesh.rotation.x += n.speed * 2;
                n.mesh.rotation.y += n.speed * 3;
            });
        }

        // Slow gentle drift of starfield
        stars.rotation.y += 0.0003;
        stars.rotation.x += 0.0001;

        accentStars.rotation.y -= 0.0002;
        accentStars.rotation.x += 0.00015;

        // Subtle twinkle via opacity pulsing
        starMat.opacity = 0.55 + 0.15 * Math.sin(t * 1.2);
        accentMat.opacity = 0.35 + 0.20 * Math.sin(t * 1.8 + 1.0);

        // Camera mouse tracking
        camera.position.x += (mouseX * 2 - camera.position.x) * 0.04;
        camera.position.y += (mouseY * 1.5 - camera.position.y) * 0.04;
        camera.lookAt(0, 0, 0);

        renderer.render(scene, camera);
    }

    animate();

    // Resize handler
    window.addEventListener('resize', () => {
        const nw = canvas.offsetWidth;
        const nh = canvas.offsetHeight;
        renderer.setSize(nw, nh);
        camera.aspect = nw / nh;
        camera.updateProjectionMatrix();
    });
}

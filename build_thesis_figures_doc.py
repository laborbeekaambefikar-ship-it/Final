"""
Generate a thesis-style DOCX documenting every figure in this repository.

Output:  ATLAS_Thesis_Figures.docx
"""

from pathlib import Path
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from PIL import Image

REPO = Path(__file__).parent

# ---------------------------------------------------------------------------
# Figure catalogue
# ---------------------------------------------------------------------------

# Each entry: (filename, chapter, fig_number, title, intro, body, image_width_in)

FIGURES = [
    # ----------------- Chapter 3 — System Design and Architecture -----------------
    {
        "file": "Overall Architecture.png",
        "chapter": 3, "num": "3.1",
        "title": "Overall System Architecture of the ATLAS Smart Warehouse AGV",
        "width_in": 6.0,
        "intro": (
            "The overall system architecture defines the high-level decomposition of the "
            "ATLAS Smart Warehouse AGV into cooperating subsystems and the data flow that "
            "binds them together. It is presented at the start of the design chapter "
            "because every subsequent design decision—mechanical, electrical, and "
            "algorithmic—is constrained by the interfaces this architecture establishes. "
            "The diagram further makes explicit the boundary between simulation-only "
            "components and the components that survive the transition to physical hardware, "
            "which is central to the deployment strategy of the project."
        ),
        "body": (
            "As shown in Fig. 3.1, the system consists of five tightly coupled architectural "
            "layers stacked from the physics layer at the bottom to the human–machine "
            "interface at the top. The lowest layer is the physics and hardware layer, which "
            "in simulation is provided by Gazebo Classic 11 for the differential-drive AGV "
            "and by Gazebo Harmonic for the UR3 manipulator subsystem. This layer publishes "
            "ground-truth pose, wheel encoder ticks, IMU readings, line-array sensor states "
            "and RFID detections, and on the actuation side it consumes wheel velocity "
            "commands and joint trajectories. Sitting directly above it is the perception "
            "layer, which fuses these heterogeneous streams into compact mission-level "
            "abstractions such as the lateral tracking error from the line-following sensor "
            "array, the heading estimate produced by IMU complementary filtering and the "
            "junction-tag identity reported by the RFID reader. "
            ""
            "The control layer above perception is responsible for low-level closed-loop "
            "behaviour. It contains the line-following PD controller that converts lateral "
            "error into an angular-velocity correction, a dedicated turn controller that "
            "manages discrete heading changes at junctions, a per-wheel velocity PID loop "
            "and a velocity arbiter that enforces single-source command authority across "
            "manual, autonomous, safety and recovery sources. The planning layer is built "
            "around the twelve-state mission Finite State Machine, the priority-aware order "
            "queue, the AGV global path planner and the MoveIt 2-based motion planner of the "
            "UR3 arm. These elements transform high-level mission requests into discrete "
            "navigation goals and time-parameterised arm trajectories. "
            ""
            "All of the above communicate through a unified ROS 2 middleware backbone, drawn "
            "explicitly in the figure as the central spine connecting every block. ROS 2's "
            "Data Distribution Service transport, quality-of-service profiles and lifecycle "
            "node management provide deterministic, decoupled and discoverable communication "
            "between processes. The topmost layer is the operator-facing GUI and HMI block, "
            "which exposes mission start, pause, reset and emergency-stop controls and "
            "displays live telemetry such as battery voltage, current state and pending "
            "order count. The architectural decomposition shown here directly supports the "
            "project's stated 70% code-reuse target: only the perception driver nodes and "
            "the actuation interface nodes need to be replaced when migrating from Gazebo "
            "to physical hardware, while the control, planning, mission and HMI layers are "
            "carried across unchanged."
        ),
    },
    {
        "file": "Hardware flow diagram.png",
        "chapter": 3, "num": "3.2",
        "title": "Hardware Architecture and Signal-Flow Diagram of the AGV",
        "width_in": 6.0,
        "intro": (
            "The hardware architecture diagram captures the electrical decomposition of the "
            "AGV platform, identifying every microcontroller, sensor, actuator and power "
            "stage as well as the directional signal flow between them. It complements the "
            "software architecture by anchoring each ROS 2 node to the physical port or bus "
            "from which its data originates. Documenting the hardware in this form is "
            "essential for reproducibility and for deriving the GPIO pin map and current "
            "budget that drive the physical build."
        ),
        "body": (
            "As shown in Fig. 3.2, the hardware architecture is organised around a "
            "Raspberry Pi 4 (4 GB) acting as the on-board ROS 2 host and a secondary "
            "low-level microcontroller, an ESP32 module, that handles real-time motor "
            "control and sensor sampling. The Raspberry Pi runs ROS 2 Humble together "
            "with the line-follower node, the velocity arbiter, the mission FSM and the "
            "GUI bridge, while the ESP32 publishes raw sensor frames over a USB-CDC link "
            "and consumes wheel velocity commands at 50 Hz. Mounted on the chassis is an "
            "eight-channel reflectance sensor array that feeds analogue line-position "
            "signals into the ESP32's ADC inputs, an MPU-6050 IMU communicating over I²C "
            "for heading estimation, and an RDM6300 125 kHz RFID reader on a UART bus that "
            "detects passive junction tags embedded in the warehouse floor. "
            ""
            "On the actuation side the ESP32 drives an L298N dual H-bridge motor driver, "
            "which delivers regulated power to two 12 V DC geared motors fitted with "
            "magnetic Hall encoders. The encoder signals close back into the ESP32's "
            "interrupt-capable GPIO pins and are used by the per-wheel velocity PID loop. "
            "A power distribution stage at the right side of the diagram shows a 12 V "
            "lithium-ion battery feeding a buck converter that produces a clean 5 V/3 A rail "
            "for the digital electronics, while the motor stage receives unregulated 12 V "
            "directly through a fuse and a manual emergency-stop switch. This separation "
            "between digital and motor power is critical for reducing brown-outs on the "
            "Raspberry Pi during high-current motor transients. "
            ""
            "The diagram also highlights the wireless interface: a Wi-Fi link from the "
            "Raspberry Pi connects the AGV to the supervisory laptop running the GUI and "
            "rviz2, while ROS 2 DDS discovery operates transparently across this link. "
            "Optional auxiliary peripherals—a buzzer for status indication and front and "
            "rear status LEDs—are wired to spare GPIOs and are toggled by the FSM in "
            "specific states such as ARM_PICK, EMERGENCY and DOCKED. This explicit "
            "bottom-up documentation of every signal path enables the simulation-to-"
            "hardware conversion table referenced later in the report and ensures that "
            "the physical build can be wired without ambiguity."
        ),
    },
    {
        "file": "Ros2 node Communication.png",
        "chapter": 3, "num": "3.3",
        "title": "ROS 2 Node Communication and Topic Graph",
        "width_in": 6.0,
        "intro": (
            "Because the entire control stack is built on ROS 2, the runtime behaviour of "
            "the system is most clearly expressed as a directed graph of nodes connected "
            "by typed topics, services and actions. The communication diagram in this "
            "section gives a process-level view of that graph, listing every node executable "
            "deployed on the AGV and on the supervisory machine and showing how messages "
            "move between them. This view is the authoritative reference for all "
            "inter-process contracts in the project."
        ),
        "body": (
            "As shown in Fig. 3.3, the system consists of approximately a dozen ROS 2 "
            "nodes distributed across the ATLAS packages and grouped into three logical "
            "clusters: sensing, control and mission. On the sensing side the line_sensor_node "
            "publishes a custom LineArray message at 50 Hz, the imu_node publishes "
            "sensor_msgs/Imu, the rfid_node publishes a JunctionTag message containing the "
            "12-byte UID of the most recently detected tag, and the odometry_node fuses "
            "wheel-encoder ticks with IMU yaw to publish nav_msgs/Odometry on /odom and the "
            "associated /tf transforms. These messages flow into the line_follower_node "
            "(running the PD controller), the turn_controller_node and the junction_detector_"
            "node, which together produce candidate geometry_msgs/Twist commands on dedicated "
            "topics such as /cmd_vel/lf and /cmd_vel/turn. "
            ""
            "The velocity_arbiter_node subscribes to all candidate command topics and to a "
            "/safety/estop boolean from the safety_node; it then publishes a single "
            "/cmd_vel topic at 20 Hz to the differential-drive plugin. The arbiter enforces "
            "the priority order safety > teleop > navigation > line-follow, ensuring that an "
            "active emergency-stop or manual joystick command can always override autonomous "
            "behaviour. Above the control cluster sits the mission_fsm_node, which exposes a "
            "set of ROS 2 actions—StartMission, PauseMission and ResetMission—and a service "
            "to enqueue Order messages with priority. The FSM consumes the JunctionTag "
            "stream, the AGV pose and arm status, and produces high-level intent messages "
            "consumed by the planner and the GUI. "
            ""
            "On the manipulation side the diagram shows MoveIt 2's move_group node, the UR3 "
            "ros2_control hardware interface and the custom mtc_pickplace node that drives "
            "the pick-and-place pipeline using the MoveIt Task Constructor. These nodes "
            "communicate using FollowJointTrajectory action goals and the standard "
            "/joint_states topic. Finally the gui_bridge_node aggregates telemetry into a "
            "single AtlasStatus message that the PyQt5-based GUI on the operator laptop "
            "consumes. By making every connection in the graph explicit, the figure also "
            "documents the QoS profiles selected for each link—reliable + transient-local "
            "for mission state, best-effort for high-rate sensor streams—which is what "
            "permits the system to operate robustly over a wireless link."
        ),
    },
    {
        "file": "Pipeline.png",
        "chapter": 3, "num": "3.4",
        "title": "End-to-End Material-Handling Mission Pipeline",
        "width_in": 6.0,
        "intro": (
            "The mission pipeline diagram complements the architecture and node graph by "
            "showing the temporal sequence of operations that the AGV and arm subsystems "
            "execute when fulfilling a single material-handling order. While the previous "
            "diagrams describe what exists, this diagram describes what happens, "
            "highlighting the hand-offs between perception, planning and manipulation. "
            "It is the conceptual backbone against which the state machine and the order "
            "manager are designed."
        ),
        "body": (
            "As shown in Fig. 3.4, the system consists of a left-to-right pipeline of "
            "stages that together implement a complete pick-up-transport-deliver cycle. "
            "The pipeline begins with order reception, in which a new request is pushed "
            "into the priority queue either by the operator GUI or by an external "
            "warehouse-management interface. The order is parsed into a structured tuple "
            "containing the source shelf identifier, the destination station, the priority "
            "level and the requested item descriptor. Once the queue head is selected, the "
            "AGV navigation stage is triggered: the global planner converts the source "
            "shelf coordinate into a sequence of waypoints aligned with the warehouse "
            "spine corridor, and the line-following PD controller keeps the chassis on the "
            "tape route while the junction detector counts and identifies RFID tags to "
            "anchor pose. "
            ""
            "When the AGV arrives at the source shelf and reaches a stable, low-velocity "
            "docking pose, the pipeline transitions to the arm pick stage. Here MoveIt 2 "
            "plans a collision-free trajectory from the home configuration to a pre-grasp "
            "pose, descends along an approach axis, closes the Robotiq 2F-85 gripper around "
            "the part and lifts to a transport-safe pose. The MoveIt Task Constructor "
            "decomposes this into modular sub-stages so that planner failures at any "
            "sub-stage can be retried without replanning the entire motion. After grasping "
            "succeeds and is confirmed by gripper-current feedback, the arm returns to a "
            "compact transport configuration that keeps the payload close to the chassis "
            "and lowers the centre of gravity. "
            ""
            "The pipeline then re-enters the navigation stage, but this time bound for the "
            "destination station. During transport the velocity arbiter caps the maximum "
            "AGV speed to a slower limit because the arm is loaded; this coupling between "
            "subsystem state and motion limits is one of the integration features that the "
            "diagram makes explicit. On arrival the arm place stage releases the part onto "
            "the destination shelf, and a delivery-confirmation message is published to "
            "the FSM, which removes the order from the queue and emits a success event to "
            "the GUI. Recovery and retry branches are drawn as dashed loops back to "
            "earlier stages, capturing failure modes such as line loss, missed junctions, "
            "grasp failure and timeout. The figure therefore serves not only as a workflow "
            "diagram but as a contract for the FSM transition table and for the recovery "
            "behaviours implemented in the velocity arbiter and the MoveIt pipeline."
        ),
    },
    {
        "file": "Agv Motion COntrol Pipeline using Pd control.png",
        "chapter": 3, "num": "3.5",
        "title": "AGV Line-Following PD Motion-Control Pipeline",
        "width_in": 6.5,
        "intro": (
            "Closed-loop line following is the mechanism by which the AGV translates a "
            "high-level navigation goal into wheel velocities, and a PD controller is used "
            "in preference to a full PID because the integral term is unnecessary for the "
            "near-zero steady-state error of a tape-following task. The motion-control "
            "pipeline diagram presented in this section exposes every transformation in the "
            "loop, from raw analogue line-array readings to per-wheel velocity setpoints. "
            "Understanding this pipeline is essential before discussing the gain tuning and "
            "tracking-accuracy results reported later."
        ),
        "body": (
            "As shown in Fig. 3.5, the system consists of a cascade of signal-processing "
            "and control blocks that together form the line-following loop. The pipeline "
            "begins with the eight-channel reflectance sensor array mounted across the "
            "front of the chassis. The raw analogue signals are read by the ESP32's ADC, "
            "thresholded against a calibrated dark/light boundary, and converted into a "
            "weighted-average lateral position estimate by the line-position-estimator "
            "block. This produces a continuous error signal e(t) expressed in millimetres "
            "of lateral offset between the chassis centre and the centre of the tape. "
            ""
            "The error signal is passed to the PD-controller block, whose output is the "
            "angular velocity correction ω_corr = K_p · e(t) + K_d · de(t)/dt, with the "
            "derivative term implemented as a first-order backward difference filtered by "
            "a single-pole low-pass to suppress sensor noise. K_p and K_d are the two free "
            "parameters of the controller; the report selects them empirically to achieve "
            "sub-5 mm lateral tracking accuracy at the cruise velocity. The base linear "
            "velocity v_lin is supplied as a separate setpoint that the FSM adjusts based "
            "on whether the chassis is in straight, junction-approach or arm-loaded mode. "
            "The (v_lin, ω_corr) pair is then fed into the differential-drive inverse-"
            "kinematics block which computes the per-wheel angular velocities ω_L and ω_R "
            "given the wheel separation L and wheel radius r. "
            ""
            "Each wheel velocity setpoint enters its own dedicated motor-velocity PID loop "
            "running at 100 Hz on the ESP32. This inner loop closes around the encoder "
            "feedback, compensating for friction, motor non-linearity, payload-induced load "
            "torque and supply-voltage sag. Its output is a PWM duty cycle clamped to a "
            "safe range and applied to the L298N H-bridge driver. The diagram further shows "
            "two parallel branches that feed back into the loop: a heading estimate from "
            "the IMU complementary filter is used to detect when the chassis has drifted "
            "off-line for longer than a fail-safe timeout, in which case the line-recovery "
            "behaviour is triggered; and a junction-detector branch compares the sensor "
            "pattern with templates for T-junctions, cross-junctions and dead-ends and "
            "emits discrete events that the FSM consumes. The pipeline thus integrates "
            "continuous PD control with discrete event detection in a way that is simple "
            "to tune yet sufficient for the structured, tape-instrumented warehouse "
            "environment targeted by the project."
        ),
    },
    {
        "file": "Priority based order management.png",
        "chapter": 3, "num": "3.6",
        "title": "Priority-Based Order-Management Workflow",
        "width_in": 5.5,
        "intro": (
            "Real warehouse operations rarely involve a single order at a time, and even "
            "in a research demonstrator the AGV must be able to receive, queue, sequence "
            "and complete multiple requests with possibly different urgency. The "
            "priority-based order-management workflow diagram defines exactly how this "
            "queueing policy is implemented and how it interacts with the mission FSM. "
            "It is therefore a key part of the planning layer described earlier in the "
            "architecture."
        ),
        "body": (
            "As shown in Fig. 3.6, the system consists of an order-input stage, a "
            "priority-queue data structure, a dispatcher and a feedback loop that closes "
            "around mission completion events. New orders enter the system from one of two "
            "sources: the operator-facing GUI, which constructs an Order message containing "
            "the source shelf, destination station, item descriptor and a priority integer "
            "drawn from the set {LOW, NORMAL, HIGH, EMERGENCY}; and an external warehouse-"
            "management interface that can publish orders programmatically through a ROS 2 "
            "service. Both sources are funnelled through an order-validation block that "
            "rejects requests with unknown shelf identifiers or with destinations that are "
            "occluded according to the latest navigation map. "
            ""
            "Validated orders are pushed into a binary-heap priority queue keyed on a "
            "composite key (priority_level, arrival_timestamp). The composite key is what "
            "guarantees the second-most important property of the queue—FIFO behaviour "
            "within a priority band—so that orders of equal priority are still served in "
            "the order they were received. A dedicated dispatcher routine pops the head of "
            "the queue whenever the FSM enters the IDLE or AWAITING_NEXT_ORDER state, and "
            "it forwards the popped tuple to the AGV path planner and the arm task "
            "selector. Because the dispatcher and the FSM are decoupled through ROS 2 "
            "topics rather than function calls, an EMERGENCY-priority order can be "
            "inserted at any time and will simply preempt the next idle event without "
            "destabilising the currently executing motion. "
            ""
            "The figure also illustrates the feedback path along which the FSM "
            "acknowledges the start of an order, reports intermediate progress through "
            "OrderStatus messages and finally publishes either an ORDER_COMPLETED or an "
            "ORDER_FAILED event. The order manager listens to these events and either "
            "removes the order from its bookkeeping store or returns it to the queue with "
            "an incremented retry counter. A maximum retry policy is enforced so that "
            "persistently failing orders are escalated to the operator instead of looping "
            "indefinitely. Beyond simple sequencing, the diagram captures one more "
            "important property: the queue is observable. Every operation on it emits a "
            "QueueDelta message that the GUI uses to render the order list in real time, "
            "giving the operator full visibility of what is pending, what is in progress "
            "and what has been completed. This combination of priority-aware dispatch, "
            "decoupled execution and live observability is what allows the AGV to behave "
            "predictably under realistic, multi-order operating conditions."
        ),
    },
    {
        "file": "12 state.png",
        "chapter": 3, "num": "3.7",
        "title": "Twelve-State Mission Finite State Machine",
        "width_in": 6.0,
        "intro": (
            "The behaviour of the AGV across an entire mission is governed by a Finite "
            "State Machine (FSM) with twelve discrete states, designed so that every "
            "operationally meaningful situation—from idling on the dock to recovering from "
            "a lost line—corresponds to exactly one state. Encoding the mission lifecycle "
            "in this explicit form supports formal analysis, simplifies debugging and is "
            "what allows the velocity arbiter, the GUI and the order manager to reason "
            "about the AGV's current intent without ambiguity."
        ),
        "body": (
            "As shown in Fig. 3.7, the system consists of twelve states, each represented "
            "by a labelled node, and a set of directed transitions, each annotated with the "
            "guard condition or event that triggers it. The IDLE state is the resting state "
            "in which the AGV is docked, motors are inhibited and the FSM is awaiting an "
            "order. From IDLE the FSM transitions to the FETCH_ORDER state as soon as the "
            "order manager publishes a non-empty queue head, copying the source and "
            "destination identifiers into local mission variables. The next state, "
            "NAVIGATE_TO_PICK, activates the line-following PD controller and the junction "
            "counter and is left only when the junction detector confirms arrival at the "
            "expected source-shelf RFID tag. "
            ""
            "On arrival the FSM enters DOCK_AT_PICK, where a low-speed precision-docking "
            "manoeuvre aligns the chassis with the shelf within a configurable tolerance. "
            "Successful docking advances the FSM to ARM_PICK, in which the MoveIt 2 "
            "pick-and-place pipeline is invoked through a ROS 2 action; the FSM blocks on "
            "the action result and inspects the gripper-current feedback to confirm a "
            "stable grasp. With a confirmed grasp the FSM transitions to "
            "NAVIGATE_TO_DROP, similar in structure to the pick navigation but with reduced "
            "maximum velocity because the arm is loaded. DOCK_AT_DROP and ARM_PLACE follow "
            "the same template as their pick counterparts and culminate in an "
            "ORDER_COMPLETED event being emitted to the order manager. "
            ""
            "Around the nominal happy-path states the diagram shows three exception "
            "states. EMERGENCY is a high-priority absorbing-style state entered whenever "
            "the safety node asserts the e-stop topic; it forces zero velocity, retracts "
            "the arm to a compact pose and waits for an explicit operator reset. "
            "LINE_LOST is entered when the line-array sensor reports an off-line condition "
            "for longer than a tunable timeout, and it executes a short reverse-and-search "
            "recovery before either re-entering navigation or escalating to EMERGENCY. "
            "ARM_FAULT is entered on MoveIt or hardware-interface errors and either retries "
            "the failed sub-stage or aborts the current order. The remaining state, "
            "RETURN_HOME, is invoked after the queue empties or after a reset and brings "
            "the chassis back to the dock. The transition table that accompanies the "
            "diagram in Appendix A enumerates every guard condition, event and side-effect, "
            "providing a single source of truth for all behavioural decisions made by the "
            "AGV over the course of an operating cycle."
        ),
    },
    {
        "file": "Arm Pose Stages.png",
        "chapter": 3, "num": "3.8",
        "title": "UR3 Manipulator Pose Sequence for Pick-and-Place",
        "width_in": 5.5,
        "intro": (
            "Whereas the AGV navigation behaviour is captured by the FSM, the arm "
            "behaviour is captured by an explicit sequence of named end-effector poses "
            "that together implement a single pick-and-place cycle. The pose-stage diagram "
            "shown in this section visualises these poses in their kinematic order and "
            "annotates each one with the role it plays in the motion plan, providing the "
            "geometric intuition that the MoveIt Task Constructor encodes algorithmically."
        ),
        "body": (
            "As shown in Fig. 3.8, the system consists of six labelled UR3 configurations "
            "rendered side by side in their natural temporal order: HOME, PRE_GRASP, "
            "GRASP, LIFT, TRANSPORT and PLACE. The HOME pose is a compact, kinematically "
            "central configuration with all joints near zero offset, chosen because it "
            "minimises the manipulator's footprint while the AGV is moving between "
            "stations. From HOME the arm extends to the PRE_GRASP pose, which is "
            "positioned a short approach distance directly above the target object; this "
            "intermediate pose decouples gross transport from fine alignment and allows "
            "MoveIt's Cartesian planner to compute a straight-line descent in the next "
            "stage. "
            ""
            "The GRASP pose places the gripper TCP exactly at the planned grasp frame, "
            "with its closing axis aligned to the principal inertia axis of the part. The "
            "Robotiq 2F-85 gripper is then commanded to close, and stable contact is "
            "confirmed by monitoring gripper current and finger position. The subsequent "
            "LIFT pose retracts the gripper vertically by a fixed clearance distance "
            "before any horizontal motion is permitted; this stage is critical because it "
            "prevents the part from catching on the shelf edge during retreat. The "
            "TRANSPORT pose moves the elbow inward and the wrist downward so that the "
            "payload is held close to the chassis centre, lowering the combined centre of "
            "gravity and reducing tip-over risk while the AGV accelerates toward the "
            "destination. "
            ""
            "Finally the PLACE stage mirrors the GRASP stage above the destination shelf: "
            "the arm executes an aligned approach, opens the gripper, retracts to a "
            "post-place clearance pose and returns to HOME, completing the cycle. The "
            "diagram makes explicit that each pose is parameterised: the approach offset, "
            "the grasp frame, the lift clearance and the transport joint vector are all "
            "configurable through the MoveIt Task Constructor's parameter interface, "
            "which means that the same sequence can be reused for parts of different "
            "geometry by changing only those numerical parameters. The figure also "
            "documents the planner that is selected for each transition; Cartesian "
            "interpolation is used between PRE_GRASP and GRASP and between LIFT and "
            "TRANSPORT, while OMPL's RRT-Connect handles HOME-to-PRE_GRASP and "
            "TRANSPORT-to-PLACE because those transitions span larger volumes of "
            "configuration space and need full collision-aware sampling. Together these "
            "design choices yield repeatable, collision-free pick-and-place cycles within "
            "the UR3's 500 mm reach radius."
        ),
    },
    # ----------------- Chapter 4 — Implementation and Simulation Environment -----------------
    {
        "file": "layour warehouse.png",
        "chapter": 4, "num": "4.1",
        "title": "Two-Dimensional Warehouse Floor Layout",
        "width_in": 6.0,
        "intro": (
            "The simulation environment used to validate the ATLAS Smart Warehouse AGV "
            "begins from a top-view two-dimensional floor layout that fixes the position "
            "of every shelf, dispatch station, pickup station and tape route. This layout "
            "is the geometric specification from which both the Gazebo world file and the "
            "Nav2 occupancy map are derived. Reproducing it here makes the spatial "
            "assumptions of the entire simulation explicit."
        ),
        "body": (
            "As shown in Fig. 4.1, the system consists of a rectangular warehouse footprint "
            "of fixed external dimensions, organised internally around a single longitudinal "
            "spine corridor with perpendicular shelf rows on either side. The spine "
            "corridor carries the primary tape route and is wide enough to allow the AGV "
            "to traverse it at full cruise speed without infringing the safety envelope of "
            "either shelf row. Each shelf row is composed of equally spaced storage units "
            "labelled S1 through Sn, and each storage unit is associated with a fixed "
            "world coordinate (x_s, y_s) that the order manager uses when dispatching pick "
            "orders. The dispatch station, where finished orders are placed for human "
            "collection, is drawn at one end of the corridor, while the docking and "
            "charging station occupies the opposite end. "
            ""
            "Embedded in the floor at the centre of every junction along the spine "
            "corridor and at every shelf entry is a circular RFID tag, drawn in the "
            "diagram as a small filled disc. These tags carry unique 12-byte identifiers "
            "that the AGV's RDM6300 reader detects as the chassis passes overhead. Because "
            "the tag positions are surveyed and stored in a YAML configuration file, every "
            "tag detection translates directly into a global pose anchor, dramatically "
            "simplifying odometry drift correction without requiring a full SLAM stack. "
            "The diagram also marks the line-following tape itself as a continuous "
            "high-contrast strip running along the centre of the spine corridor and "
            "branching into each shelf entry. "
            ""
            "Around the periphery of the warehouse the layout indicates fixed obstacles "
            "such as pillars and storage cabinets that the navigation map must inflate, "
            "and it also reserves a narrow human-only walkway separated from the AGV "
            "route. This separation reflects the design decision to keep human-AGV "
            "interaction zones spatially isolated, which in turn relaxes the perception "
            "requirements on the AGV (since unexpected pedestrians are unlikely on the "
            "tape route itself). The layout therefore captures the three categories of "
            "spatial information needed by the rest of the system: the tape topology used "
            "by the line follower, the discrete junction graph used by the FSM and the "
            "dispatcher, and the obstacle map used by the global planner. Every "
            "subsequent simulation figure can be interpreted as a particular rendering of "
            "this same floor specification, with progressively more visual detail layered "
            "on top."
        ),
    },
    {
        "file": "Warehouse nav map.png",
        "chapter": 4, "num": "4.2",
        "title": "Two-Dimensional Navigation and Occupancy Map",
        "width_in": 6.0,
        "intro": (
            "The two-dimensional occupancy map is the representation of the warehouse "
            "consumed by the navigation stack and by the AGV's global path planner. It is "
            "derived from the floor layout by rasterising every static obstacle into a "
            "fixed-resolution grid and inflating it according to the chassis footprint. "
            "Including this map alongside the layout makes it possible to relate planned "
            "paths back to the underlying physical geometry."
        ),
        "body": (
            "As shown in Fig. 4.2, the system consists of a grid map at a chosen "
            "resolution of 0.05 m per cell, in which free cells are rendered light, "
            "occupied cells are rendered dark and inflation cells (cells inside the "
            "robot's safety radius around an obstacle) are rendered in an intermediate "
            "shade. The map covers the full extent of the warehouse layout, with the "
            "spine corridor appearing as a continuous wide free strip flanked by long "
            "lines of occupied cells corresponding to shelf walls. The dispatch and "
            "docking areas are visible as larger free regions at either end of the "
            "corridor, and the walls of the warehouse itself appear as the outermost "
            "occupied rectangle. "
            ""
            "The map also encodes the discrete junction graph implicitly through the "
            "geometry of the corridor: every place where a perpendicular shelf entry "
            "branches off from the spine corridor produces a recognisable T-shape in the "
            "free space. When the global planner overlays its computed path, drawn in the "
            "figure as a solid coloured polyline, the path visibly hugs the centre of the "
            "spine corridor exactly along the tape route, which is the desired behaviour "
            "because the tape coincides with the local minimum of the inflated cost field. "
            "This consistency is what allows the global planner and the line-following PD "
            "controller to operate as complementary rather than conflicting layers: the "
            "global planner produces high-level waypoints anchored to RFID tags, while the "
            "PD controller closes the local loop on the tape itself. "
            ""
            "Beyond its role for navigation, the occupancy map also drives several other "
            "components of the system. The MoveIt 2 planning scene receives a small extruded "
            "version of the same map as a static collision object, ensuring that the arm "
            "treats the warehouse walls as obstacles when it plans large reconfiguration "
            "motions. The GUI uses the map as the background layer onto which it overlays "
            "the live AGV pose, the order list and the active path. And the simulation "
            "harness uses the map to verify, before launching Gazebo, that every shelf "
            "coordinate referenced by the order generator falls inside the free space and "
            "is therefore reachable. The figure thus serves both as a debugging aid and as "
            "a contract that ties the higher-fidelity 3-D Gazebo world, the abstract "
            "geometric layout and the planning algorithms together into a single coherent "
            "spatial representation."
        ),
    },
    {
        "file": "warehouse.png",
        "chapter": 4, "num": "4.3",
        "title": "Three-Dimensional Gazebo Warehouse Environment",
        "width_in": 6.0,
        "intro": (
            "Once the layout and the navigation map are fixed, they are realised as a "
            "fully three-dimensional Gazebo world that hosts the rigid-body physics, "
            "lighting, sensor models and visual meshes used during simulation. The "
            "rendered Gazebo warehouse is what the perception stack actually sees and is "
            "therefore the reference environment for every figure that follows in this "
            "chapter. Documenting it explicitly is essential for reproducibility of the "
            "reported experiments."
        ),
        "body": (
            "As shown in Fig. 4.3, the system consists of a Gazebo Classic 11 world file "
            "that instantiates the warehouse hull as a textured concrete-floor box with "
            "matte interior walls, populated by parametric shelf models arranged in two "
            "parallel rows along the spine corridor. Each shelf model is a multi-tier "
            "rigid body assembly of metal uprights and wooden trays, with collision "
            "geometry simplified to a stack of axis-aligned bounding boxes for "
            "computational efficiency while keeping the visual mesh photo-realistic. The "
            "tape route is rendered as a thin high-contrast strip texture-mapped onto the "
            "concrete floor, exactly along the centre of the spine corridor and into each "
            "shelf entry, providing the visual signal that the line-array sensor model "
            "samples. "
            ""
            "Lighting is provided by a directional sun-light approximating overhead "
            "industrial fixtures, plus point lights placed at regular intervals along the "
            "ceiling to ensure that the line-following sensor receives a usable contrast "
            "ratio everywhere on the tape. Several boxed parts of varying mass and "
            "dimension are pre-spawned on selected shelf trays so that the arm subsystem "
            "has realistic targets to grasp; their masses and friction coefficients are "
            "tuned to be reachable by the Robotiq 2F-85's gripping force without slip. "
            "Around the periphery the world includes static decorative obstacles—pallets, "
            "a forklift mesh, fire-extinguisher posts—whose collision boxes match the "
            "occupancy map's outermost obstacles so that perception, planning and physics "
            "are mutually consistent. "
            ""
            "The simulation also instantiates a number of Gazebo-specific plugins. The "
            "differential-drive plugin is attached to the AGV chassis and exposes the "
            "/cmd_vel topic that the velocity arbiter publishes on; corresponding "
            "ground-truth pose, IMU and wheel-encoder plugins expose the sensor topics "
            "consumed by the perception layer. A custom RFID plugin reads the surveyed tag "
            "positions from the same YAML file used by the layout and emits a JunctionTag "
            "message whenever the AGV's reader frame is within reading range of a tag. "
            "Time is advanced at real-time factor 1.0 on a moderately equipped laptop, "
            "though headless mode and step-mode are both supported for headless CI runs. "
            "By documenting the world contents to this level of detail, the figure makes "
            "the simulation environment a fully specified experimental platform rather "
            "than an opaque visual prop."
        ),
    },
    {
        "file": "warehouse top shot.png",
        "chapter": 4, "num": "4.4",
        "title": "Top-Down Render of the Simulated Warehouse",
        "width_in": 6.0,
        "intro": (
            "A top-down rendering of the same Gazebo world provides a useful intermediate "
            "between the abstract two-dimensional layout and the immersive perspective "
            "view, because it preserves three-dimensional shading and lighting cues while "
            "exposing the global structure of the environment. This view is included so "
            "that the reader can verify visually that the simulated warehouse matches the "
            "specification given by the layout and the navigation map."
        ),
        "body": (
            "As shown in Fig. 4.4, the system consists of an orthographic-style overhead "
            "camera that captures the entire warehouse footprint in a single frame, "
            "revealing the symmetry of the two shelf rows around the central spine "
            "corridor and the placement of the dispatch and docking areas at opposite "
            "ends. The continuity of the line-following tape is clearly visible from this "
            "perspective: it runs as a single uninterrupted strip down the corridor and "
            "branches cleanly into each shelf entry, confirming that the texture mapping "
            "in the world file is geometrically consistent with the layout drawing. "
            ""
            "The top-down render also exposes secondary structural details that are easy "
            "to overlook from a perspective camera. The aisles between successive shelves "
            "in each row are wide enough to accommodate small inspection robots but are "
            "deliberately narrower than the spine corridor so that the AGV's global "
            "planner has no incentive to enter them; this asymmetry encodes the "
            "operational discipline that the AGV traverses only the spine and the "
            "explicit shelf entries. The dispatch area at one end of the corridor is "
            "rendered as an open work surface adjacent to a workshop door, while the "
            "docking station at the opposite end shows a marked rectangular footprint "
            "that the precision-docking behaviour aligns with at the end of every "
            "mission. "
            ""
            "Sensor placement and lighting are equally informative from this view. Because "
            "the overhead lights are visible as bright spots evenly distributed along the "
            "corridor, it is easy to confirm that the line-array sensor will see uniform "
            "tape contrast throughout, eliminating one common source of line-loss events "
            "in tape-following systems. The shadows of the shelves on the floor "
            "demonstrate that the directional sun-light is configured to come from a "
            "consistent direction, which simplifies any future computer-vision extensions "
            "that may rely on shading cues. Finally, the top-down view is the natural "
            "reference frame for overlaying the global path produced by the planner, the "
            "live AGV pose and the priority-ordered list of pending orders; this is the "
            "exact composition used by the operator-facing GUI, which means that the "
            "visual model the operator works with is precisely the one shown in this "
            "figure. The top-down render therefore acts as the bridge between the "
            "abstract spatial representations introduced earlier and the immersive "
            "experiential views that follow."
        ),
    },
    {
        "file": "warehouse with agv.png",
        "chapter": 4, "num": "4.5",
        "title": "ATLAS AGV Operating Inside the Simulated Warehouse",
        "width_in": 6.0,
        "intro": (
            "Showing the AGV inside the simulated warehouse confirms that the chassis "
            "model, sensor suite and motion behaviour integrate correctly with the world "
            "file and that the visual scale matches the layout specification. This "
            "perspective view also makes it possible to inspect tape contact, sensor "
            "footprint and chassis clearance, none of which are observable from the "
            "abstract maps alone."
        ),
        "body": (
            "As shown in Fig. 4.5, the system consists of the differential-drive AGV "
            "chassis instantiated in the spine corridor of the Gazebo warehouse, with its "
            "front line-array sensor straddling the high-contrast tape and its IMU and "
            "RFID reader mounted along the chassis centreline. The chassis URDF model "
            "carries every visual feature of the physical platform: the two rear-driven "
            "wheels with realistic radius and width, a free-rolling caster at the front, "
            "the protective fenders that house the line-array, the on-board electronics "
            "enclosure and the upper deck that supports the UR3 arm base. Because the "
            "URDF is reused unchanged on the physical platform, this view is also the "
            "visual reference for the real-world build. "
            ""
            "From this perspective the relationship between the AGV and the tape route is "
            "particularly clear. The line-array sensor is positioned a few centimetres "
            "ahead of the wheel axle so that lateral error is detected slightly before the "
            "chassis itself deviates, providing the look-ahead that the PD controller "
            "needs to remain stable at cruise speed. The chassis footprint sits well "
            "inside the spine-corridor free space, with margins on either side that the "
            "occupancy-map inflation step has reserved as a safety buffer. The colour and "
            "geometry of the wheels are identical to those used in the inverse-kinematics "
            "block, so the differential-drive plugin's response to /cmd_vel reflects the "
            "kinematic equations of the report rather than an arbitrary tuned model. "
            ""
            "The view further demonstrates the sensor coverage of the platform. The IMU "
            "is mounted near the geometric centre of the chassis, minimising centripetal "
            "coupling on the yaw measurement during tight turns, while the RFID reader "
            "frame is mounted on the underside near the tape line so that its reading "
            "field intercepts every junction tag as the chassis passes overhead. The "
            "computing enclosure on top of the deck represents the Raspberry Pi and "
            "ESP32 boards in the physical build and is sized to leave clearance for the "
            "UR3 arm's swept volume during transport. Watching the AGV move through this "
            "world during a mission also confirms the dynamic behaviour of the FSM: it "
            "can be observed entering and leaving each state at the correct geometric "
            "position, validating that the integration of perception, planning and "
            "control behaves as designed. The figure thereby ties together every concept "
            "introduced earlier in the chapter—mechanical, sensory, computational and "
            "behavioural—into a single observable scenario."
        ),
    },
    {
        "file": "robotic arm.png",
        "chapter": 4, "num": "4.6",
        "title": "UR3 Robotic Arm with Robotiq Gripper in Simulation",
        "width_in": 6.0,
        "intro": (
            "The manipulation subsystem is built around a Universal Robots UR3 6-DOF "
            "collaborative arm equipped with a Robotiq 2F-85 two-finger gripper, both "
            "modelled in URDF and simulated under Gazebo Harmonic. Examining the simulated "
            "arm in isolation makes it possible to verify joint limits, link visuals, "
            "tool-frame placement and gripper geometry before they are coupled with the "
            "AGV chassis."
        ),
        "body": (
            "As shown in Fig. 4.6, the system consists of the UR3 manipulator mounted on "
            "a fixed base plate, with the Robotiq 2F-85 gripper attached at the wrist "
            "flange and a tool-frame coordinate triad rendered at the end-effector for "
            "reference. The UR3 model uses the official Universal Robots URDF, modified "
            "only by the addition of the gripper end-effector and the project's tool-"
            "centre-point transform. The six revolute joints are visible at the shoulder "
            "pan, shoulder lift, elbow, wrist 1, wrist 2 and wrist 3 positions, and each "
            "joint is annotated in the URDF with its mechanical limit, maximum velocity "
            "and effort, taken from the manufacturer's datasheet. "
            ""
            "Kinematically the arm follows the modified Denavit–Hartenberg convention used "
            "by ROS Industrial; the four DH parameters per link (a, alpha, d, theta) are "
            "tabulated in the report and form the basis of every analytical and numerical "
            "inverse kinematics computation. MoveIt 2 is configured to use this DH model "
            "directly through the ur_description and ur_moveit_config packages, which "
            "means that the planning scene shown in the figure is not a custom "
            "approximation but the same kinematic representation used during deployment. "
            "The Robotiq 2F-85 gripper is added with a parallel-jaw URDF whose two finger "
            "joints are coupled by a mimic relationship so that a single command opens or "
            "closes both fingers symmetrically; this matches the behaviour of the physical "
            "gripper's controller. "
            ""
            "Several practical features of the model are visible in the rendered view. "
            "The base plate is sized to match the AGV's upper deck so that the arm can be "
            "transferred from the standalone test scene to the integrated AGV scene "
            "without any pose-frame edits. The visual meshes are detailed enough to "
            "support meaningful camera-based perception experiments in future work, while "
            "the collision meshes are simplified primitives that keep planning queries "
            "fast enough for interactive use. The wrist-mounted force-torque interface "
            "specified in the URDF makes it possible to extend the system later with "
            "force-controlled insertion behaviours without re-engineering the model. "
            "Finally, the tool-frame triad at the end-effector is placed exactly between "
            "the gripper fingers' contact surfaces, which is the convention assumed by "
            "the MoveIt Task Constructor when planning approach and retreat motions. The "
            "figure therefore documents not only the visual appearance of the arm but the "
            "dimensional and kinematic conventions on which all higher-level manipulation "
            "behaviour is built."
        ),
    },
    {
        "file": "image-removebg-preview.png",
        "chapter": 4, "num": "4.7",
        "title": "Physical ATLAS AGV Chassis (Real-World Build)",
        "width_in": 4.5,
        "intro": (
            "Although the bulk of validation is performed in simulation, the project also "
            "includes a physical realisation of the AGV chassis intended to demonstrate "
            "the simulation-to-hardware portability of the software stack. The "
            "background-removed photograph reproduced in this section captures that real "
            "chassis in isolation so that its mechanical layout can be compared "
            "side-by-side with the simulated URDF model presented earlier."
        ),
        "body": (
            "As shown in Fig. 4.7, the system consists of the as-built ATLAS chassis "
            "imaged with the surrounding workspace removed for clarity. The platform is a "
            "low, flat differential-drive robot whose external dimensions match the URDF "
            "model used in Gazebo, with the two rear-driven wheels visible at the back "
            "edge and the free caster at the front. Mounted on the upper deck are the "
            "Raspberry Pi 4 single-board computer and the ESP32 microcontroller, "
            "interconnected by a short USB-CDC cable that carries the high-rate sensor "
            "stream and the wheel-velocity command stream defined in the hardware "
            "architecture diagram. The L298N motor driver and the protective fuse-and-"
            "switch module sit between the battery pack and the motors, isolating motor "
            "transients from the digital electronics. "
            ""
            "Underneath the front of the chassis is the eight-channel reflectance line-"
            "array sensor, mounted exactly at the simulated stand-off height so that the "
            "calibrated dark/light threshold transfers directly between simulation and "
            "hardware. The MPU-6050 IMU is fastened to the centre of the deck, oriented "
            "with its X-axis pointing forward to match the URDF convention, and the "
            "RDM6300 RFID reader is positioned on the underside near the tape-following "
            "line so that its reading field intercepts the surveyed junction tags. Two "
            "status LEDs and a small piezo buzzer are wired to spare GPIOs of the ESP32; "
            "the FSM toggles them in distinguished states to give the operator a low-"
            "bandwidth view of the AGV's behaviour even without a network connection to "
            "the GUI. "
            ""
            "Because the chassis was specifically engineered to be a one-to-one physical "
            "realisation of the simulated URDF, the figure also serves as evidence for "
            "the project's headline portability claim. Each component visible in the "
            "photograph has a counterpart in the simulation: the differential-drive "
            "Gazebo plugin substitutes for the real motors and encoders, the Gazebo IMU "
            "plugin substitutes for the MPU-6050, the custom RFID Gazebo plugin "
            "substitutes for the RDM6300, and the simulated reflectance sensor "
            "substitutes for the line-array board. With the perception drivers and "
            "actuator interfaces being the only nodes that change between the two "
            "deployments, the bulk of the ROS 2 stack—line follower, velocity arbiter, "
            "FSM, order manager and GUI—runs unchanged. The photograph therefore closes "
            "the loop between the abstract architectural diagrams of Chapter 3, the "
            "simulation views earlier in this chapter and the real hardware on which "
            "future field experiments will be conducted."
        ),
    },
]


# ---------------------------------------------------------------------------
# Document construction helpers
# ---------------------------------------------------------------------------

def set_a4_and_margins(doc):
    section = doc.sections[0]
    # A4
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    # 1-inch margins
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(1.0)


def set_default_font(doc):
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), "Times New Roman")


def add_justified_paragraph(doc, text, *, first_line_indent_in=0.0, space_after_pt=8):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf = p.paragraph_format
    pf.space_after = Pt(space_after_pt)
    pf.line_spacing = 1.15
    if first_line_indent_in:
        pf.first_line_indent = Inches(first_line_indent_in)
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(12)
    return p


def add_centered_image(doc, path, width_in):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run()
    run.add_picture(str(path), width=Inches(width_in))
    return p


def add_figure_caption(doc, num, title):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(14)
    run = p.add_run(f"Figure {num}: {title}")
    run.bold = True
    run.font.name = "Times New Roman"
    run.font.size = Pt(12)
    return p


def add_chapter_heading(doc, number, title):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(18)
    run = p.add_run(f"Chapter {number}\n{title}")
    run.bold = True
    run.font.name = "Times New Roman"
    run.font.size = Pt(18)
    return p


def add_section_heading(doc, number, title):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(8)
    run = p.add_run(f"{number}  {title}")
    run.bold = True
    run.font.name = "Times New Roman"
    run.font.size = Pt(14)
    return p


def add_page_break(doc):
    from docx.enum.text import WD_BREAK
    p = doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)


# ---------------------------------------------------------------------------
# Build the document
# ---------------------------------------------------------------------------

def build():
    doc = Document()
    set_a4_and_margins(doc)
    set_default_font(doc)

    # ---- Title page ----
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_before = Pt(140)
    run = title_p.add_run(
        "Development of a Smart Autonomous Guided Vehicle\n"
        "Based Warehouse System with Integrated Robotic Arm\n"
        "for Material Picking and Transfer"
    )
    run.bold = True
    run.font.name = "Times New Roman"
    run.font.size = Pt(20)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_p.paragraph_format.space_before = Pt(36)
    sub_run = sub_p.add_run("Figure Documentation Supplement")
    sub_run.italic = True
    sub_run.font.name = "Times New Roman"
    sub_run.font.size = Pt(16)

    info_p = doc.add_paragraph()
    info_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    info_p.paragraph_format.space_before = Pt(60)
    info_run = info_p.add_run(
        "Sheikh Noorul A. Usmani  (22AEB180)\n"
        "Arbaz Rashid  (22AEB485)\n\n"
        "Under the guidance of\n"
        "Prof. Mohammad Muzammil\n\n"
        "Zakir Hussain College of Engineering & Technology\n"
        "Aligarh Muslim University\n"
        "2025–26"
    )
    info_run.font.name = "Times New Roman"
    info_run.font.size = Pt(13)

    add_page_break(doc)

    # ---- Group figures by chapter ----
    by_chapter = {}
    for f in FIGURES:
        by_chapter.setdefault(f["chapter"], []).append(f)

    chapter_titles = {
        3: "System Design and Architecture",
        4: "Implementation and Simulation Environment",
    }

    chapter_intros = {
        3: (
            "This chapter presents the complete design of the ATLAS Smart Warehouse "
            "AGV, beginning with the high-level architecture that decomposes the system "
            "into cooperating layers and progressing through the hardware-flow, ROS 2 "
            "communication, mission-pipeline, motion-control, order-management, "
            "state-machine and manipulator-pose representations. Each figure in this "
            "chapter captures a single, well-bounded design concern, and together they "
            "form the complete specification from which the implementation is derived."
        ),
        4: (
            "This chapter describes how the design specified in Chapter 3 is realised in "
            "a fully working simulation environment, and how that environment relates to "
            "the physical hardware build. The figures in this chapter progress from the "
            "abstract two-dimensional warehouse layout through the navigation map, the "
            "three-dimensional Gazebo world, the AGV in operation and the UR3 arm in "
            "isolation, finishing with a photograph of the as-built AGV chassis to "
            "demonstrate the simulation-to-hardware portability claim of the project."
        ),
    }

    for ch in sorted(by_chapter.keys()):
        add_chapter_heading(doc, ch, chapter_titles[ch])
        add_justified_paragraph(doc, chapter_intros[ch], first_line_indent_in=0.3)

        for fig in by_chapter[ch]:
            # Section heading per figure
            add_section_heading(doc, fig["num"], fig["title"])
            # Intro paragraph
            add_justified_paragraph(doc, fig["intro"], first_line_indent_in=0.3)
            # Image
            img_path = REPO / fig["file"]
            assert img_path.exists(), f"missing image: {img_path}"
            add_centered_image(doc, img_path, fig["width_in"])
            # Caption
            add_figure_caption(doc, fig["num"], fig["title"])
            # Body paragraph
            add_justified_paragraph(doc, fig["body"], first_line_indent_in=0.3)
            # Page break between figures so each has its own page
            add_page_break(doc)

        # Don't add an extra page break at end of chapter (the last fig already added one)

    out_path = REPO / "ATLAS_Thesis_Figures.docx"
    doc.save(out_path)
    return out_path


if __name__ == "__main__":
    out = build()
    size = out.stat().st_size
    print(f"Wrote {out}  ({size:,} bytes)")

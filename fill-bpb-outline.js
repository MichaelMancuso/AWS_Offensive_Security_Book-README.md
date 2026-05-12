/*
 * fill-bpb-outline.js
 *
 * Generates BPB_Book_Outline_FILLED.docx — the BPB Book Outline Template,
 * populated from the AWS Offensive Security manuscript.
 *
 * Usage:
 *   npm install        # only needed once; uses the same node_modules as build-docx.js
 *   node fill-bpb-outline.js
 *
 * Output:
 *   BPB_Book_Outline_FILLED.docx (next to this script)
 *
 * The original "BPB Book Outline Template.docx" is never modified — this
 * script writes a fresh file you can hand straight to the publisher.
 */

'use strict';

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, LevelFormat,
  BorderStyle, WidthType, ShadingType, VerticalAlign, PageNumber,
  PageBreak, TabStopType, TabStopPosition,
} = require('docx');

// ----------------------------------------------------------------------------
// Style constants (matched 1:1 to the BPB template's typography)
//
//   Font:       Times New Roman everywhere
//   Body:       14pt
//   Italic hint: 12pt italic
//   Pink standalone heading (Book Description, Chapter X, etc.):  14pt bold
//   Pink inline label (Book Title –, Subtitle -, Who is this book for -): 16pt bold
//   "Outline Template" page title (centered): 24pt underlined
//   "Section I" / "Section II" centered:      20pt bold underlined
//   Bullet topics: 12pt regular
//
// docx-js sizes are in HALF-POINTS, so 14pt = 28, 16pt = 32, 20pt = 40, 24pt = 48.
// ----------------------------------------------------------------------------
const FONT = 'Times New Roman';
const SIZE_BODY        = 28;   // 14pt
const SIZE_HINT        = 24;   // 12pt italic
const SIZE_PINK_HEAD   = 28;   // 14pt bold
const SIZE_INLINE_LBL  = 32;   // 16pt bold
const SIZE_PAGE_TITLE  = 48;   // 24pt underlined
const SIZE_SECTION     = 40;   // 20pt bold underlined
const SIZE_BULLET      = 24;   // 12pt
const SIZE_NOTE        = 20;   // 10pt italic (small char-count notes)

const BPB_PINK   = '8E2A6E';   // pink color in the original template
const BODY_INK   = '111111';
const ITALIC_HINT = '7A7A7A';
const TABLE_HEADER_BG = 'EFE6EE';
const BORDER = { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC' };
const ALL_BORDERS = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER };

// US Letter, 1in margins -> 9360 DXA content width
const PAGE_W = 12240;
const PAGE_H = 15840;
const MARG   = 1440;
const CONTENT_W = PAGE_W - 2 * MARG;

// ----------------------------------------------------------------------------
// Helpers
// ----------------------------------------------------------------------------
function P(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120, ...(opts.spacing || {}) },
    alignment: opts.alignment,
    children: [new TextRun({
      text,
      bold: opts.bold,
      italics: opts.italics,
      size: opts.size || SIZE_BODY,    // 14pt body default
      color: opts.color || BODY_INK,
      font: opts.font || FONT,
    })],
  });
}

// Standalone pink heading: 14pt bold (Book Description, Key Features,
// Competitor Analysis, Tech List, Author Bio, Chapter Details, Chapter X)
function pinkHeading(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 280, after: 140 },
    children: [new TextRun({
      text,
      bold: true,
      size: opts.size || SIZE_PINK_HEAD,
      color: BPB_PINK,
      font: FONT,
    })],
  });
}

// Italic prompt/hint text in the template — 12pt italic
function italicHint(text) {
  return new Paragraph({
    spacing: { after: 80 },
    children: [new TextRun({
      text,
      italics: true,
      size: SIZE_HINT,
      color: ITALIC_HINT,
      font: FONT,
    })],
  });
}

// Topic bullets: 12pt regular
function bullet(text) {
  return new Paragraph({
    numbering: { reference: 'bpb-bullets', level: 0 },
    spacing: { after: 60 },
    children: [new TextRun({ text, size: SIZE_BULLET, font: FONT, color: BODY_INK })],
  });
}

// Page title ("Outline Template"): 24pt underlined, centered, NOT bold
function pageTitle(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 240 },
    children: [new TextRun({
      text,
      underline: {},
      size: SIZE_PAGE_TITLE,
      color: BPB_PINK,
      font: FONT,
    })],
  });
}

// Section title ("Section I" / "Section II"): 20pt bold underlined, centered
function sectionTitle(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 240, after: 240 },
    children: [new TextRun({
      text,
      bold: true,
      underline: {},
      size: SIZE_SECTION,
      color: BPB_PINK,
      font: FONT,
    })],
  });
}

// Tiny 10pt italic note used for character/word-count hints
function smallNote(text) {
  return new Paragraph({
    spacing: { after: 60 },
    children: [new TextRun({
      text, italics: true, size: SIZE_NOTE, color: ITALIC_HINT, font: FONT,
    })],
  });
}

// Pink inline label with its filled-in value: 16pt bold label + 16pt regular value
function inlineLabelValue(label, value) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [
      new TextRun({ text: label, bold: true, size: SIZE_INLINE_LBL, color: BPB_PINK, font: FONT }),
      new TextRun({ text: ' – ',           size: SIZE_INLINE_LBL, color: BPB_PINK, font: FONT }),
      new TextRun({ text: value,           size: SIZE_INLINE_LBL, color: BODY_INK, font: FONT }),
    ],
  });
}

function tableCell(text, opts = {}) {
  return new TableCell({
    borders: ALL_BORDERS,
    width: { size: opts.width, type: WidthType.DXA },
    shading: opts.fill
      ? { fill: opts.fill, type: ShadingType.CLEAR, color: 'auto' }
      : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT,
      children: [new TextRun({
        text,
        bold: opts.bold,
        italics: opts.italics,
        size: opts.size || SIZE_BULLET,       // 12pt for table cells (matches template body in tables)
        color: opts.color || BODY_INK,
        font: FONT,
      })],
    })],
  });
}

// ----------------------------------------------------------------------------
// Content
// ----------------------------------------------------------------------------

const subtitle =
  "A Red-Team Practitioner's Handbook for Amazon Web Services"; // 58 chars, under the 70-char limit

const audience =
  "This book is a practitioner's handbook for red-team operators, " +
  "offensive security consultants, detection engineers, and cloud platform " +
  "defenders who need a working model of how attacks actually land in " +
  "Amazon Web Services — not in theory, not in a curriculum, but on live " +
  "accounts, under time pressure, against active defenders.";

const description = [
  "Amazon Web Services has become the operating system of the modern enterprise, " +
  "and IAM has become its kernel. Whoever controls IAM controls the cloud, and " +
  "yet IAM is the single most-misunderstood AWS subsystem in production today. " +
  "AWS Offensive Security is the practitioner's handbook for the red-teamers, " +
  "offensive consultants, and cloud security engineers who are paid to find out " +
  "how badly that misunderstanding can be weaponized.",

  "Across twenty chapters and six parts, the book moves from foundations — the " +
  "shared-responsibility model as an attack map, the AWS attack surface, lab " +
  "construction — through reconnaissance and initial access, into the theoretical " +
  "spine of IAM exploitation: thirty-plus privilege-escalation primitives, role " +
  "chaining, cross-account abuse, and the confused-deputy landscape. From there it " +
  "shifts to service-specific tradecraft (compute, storage, data services, " +
  "containers, serverless), then post-exploitation (persistence, defense evasion, " +
  "exfiltration), CI/CD supply-chain attacks, and the practitioner's deliverables " +
  "that turn an engagement into client value. Every primitive is paired with the " +
  "CloudTrail evidence trail and the detection rule that catches it.",

  "Readers complete the book with a working operator's mental model of AWS, a " +
  "hands-on lab they can detonate against, an end-to-end privilege-escalation " +
  "matrix, MITRE ATT&CK for Cloud mappings, and the engagement deliverables — " +
  "findings, detection queries, Terraform remediation — that separate a serious " +
  "AWS red-team practice from a checkbox audit.",
];

const keyFeatures = [
  "30+ IAM privilege-escalation primitives with API and detection paths",
  "End-to-end coverage from recon to report across 20 service chapters",
  "Hands-on AWS lab harness using SCPs, CloudTrail, and Terraform",
  "MITRE ATT&CK for Cloud mapping with CloudTrail evidence per finding",
  "Operator-grade tooling: CloudFox, Pacu, PMapper, ScoutSuite, Prowler",
];

const competitorAnalysis = [
  // Each item: title, author, publisher/year, comparison
  {
    title: "AWS Penetration Testing",
    author: "Jonathan Helmus, Packt Publishing (2020)",
    note:
      "Helmus's book is a competent introduction to scanning AWS-hosted " +
      "infrastructure with conventional pentest tooling (nmap, Metasploit, " +
      "Nessus). It treats AWS as a hosting target. By contrast, this book " +
      "treats AWS itself as the attack surface — IAM, STS, service trusts, " +
      "and CloudTrail — which is where the modern threat actually lives.",
  },
  {
    title: "Hands-On AWS Penetration Testing with Kali Linux",
    author: "Karl Gilbert and Benjamin Caudill, Packt Publishing (2019)",
    note:
      "An influential early text and the original home of the Pacu framework. " +
      "Now five years out of date: predates IAM Identity Center, the 2022 " +
      "confused-deputy behavior change, IRSA / Pod Identity, IMDSv2 enforcement, " +
      "OIDC-based CI/CD trusts, and the post-2023 GuardDuty / Access Analyzer " +
      "controls. AWS Offensive Security is structured to give the operator a " +
      "current playbook against AWS as it exists in 2026.",
  },
];
const competitorPositioning =
  "Comparable courses include the SANS SEC588 (Cloud Penetration Testing) and " +
  "the OffSec OSWA / Cloud-track content. Those programs focus on a fixed " +
  "curriculum and a graded exam. This book is the operator's reference text " +
  "those programs leave out — the field manual a working AWS red-teamer keeps " +
  "open on a second monitor during an engagement.";

const techList = [
  "AWS services: IAM, STS, Organizations, IAM Identity Center (formerly AWS SSO), " +
  "EC2, SSM, Lambda, S3, EBS, EFS, RDS, Aurora, DynamoDB, Redshift, Athena, ECS, " +
  "EKS, Fargate, ECR, EventBridge, API Gateway, SNS, SQS, Step Functions, KMS, " +
  "Secrets Manager, Systems Manager Parameter Store, CloudFormation, Glue, " +
  "CodeBuild, CodePipeline, CloudTrail, GuardDuty, Access Analyzer, RAM, " +
  "VPC peering, AWS Config.",

  "Open-source offensive tooling: aws-cli v2, CloudFox, Pacu, ScoutSuite, " +
  "Prowler, CloudSploit, PMapper, quiet-riot, weirdAAL, IAM Vulnerable, " +
  "CloudGoat, enumerate-iam, leakix, GitHub secret-scanner, trufflehog.",

  "Defensive / blue-team tooling referenced: GuardDuty, Access Analyzer, " +
  "AWS Detective, AWS Config Rules, Athena (for CloudTrail forensics), " +
  "Sigma rules, MITRE ATT&CK for Cloud Navigator.",

  "Programming and infrastructure: Python 3.11+, Bash, PowerShell, Terraform " +
  "(HCL), Kali Linux, GitHub Actions (OIDC), Docker.",
];

// Author bio — three paragraphs, written to match the manuscript's voice
// (operator-grade, current, offense-and-defense). Edit freely; rerun to
// regenerate the .docx.
const authorBio = [
  "Michael Mancuso is a Senior Penetration Tester at Inovalon with over a " +
  "decade of offensive-security experience across regulated environments — " +
  "PCI, HIPAA, and SOX — where the consequences of a missed finding are " +
  "operational, not academic. His engagements span web applications, " +
  "Microsoft Active Directory, APIs, and the AWS control plane, with a " +
  "working specialty in cloud red-team tradecraft: IAM privilege escalation, " +
  "role-chain mapping, service-specific exploitation, and the CloudTrail " +
  "evidence trail every primitive leaves behind.",

  "He works the offense and the defense together on every engagement. The " +
  "findings he writes ship with the detection rule that would have caught " +
  "them and the Terraform that closes the gap, because a finding without a " +
  "hand-off to the blue team is a finding that resurfaces a quarter later. " +
  "His daily toolkit blends the standards — CloudFox, Pacu, PMapper, " +
  "ScoutSuite, Prowler — with custom tooling written in Python, Bash, and " +
  "PowerShell, the kind of harness that survives the gap between what " +
  "automated scanners catch and what an active defender actually sees.",

  "AWS Offensive Security is the field manual he wishes had existed when he " +
  "started — the operator's reference for working AWS as it exists in 2026, " +
  "written by a practitioner currently in the engagement seat.",
];

// ----------------------------------------------------------------------------
// Chapter data
// ----------------------------------------------------------------------------
const PARTS = [
  {
    section: 'Part I — Foundations',
    chapters: [
      {
        n: 1,
        title: 'The AWS Threat Model and Shared Responsibility',
        pages: 14,
        description:
          "Map the AWS shared-responsibility model onto an adversary's mental " +
          "model. This chapter identifies which layers the operator has authority " +
          "to test, how the control plane and data plane differ as attack " +
          "surfaces, and how blast radius scales from service to account to " +
          "organization. It introduces the canonical AWS attack lifecycle and " +
          "closes with the regulatory and contractual context that shapes every " +
          "engagement scope. By the end the reader can articulate exactly which " +
          "primitives are in scope, which are out of scope, and which sit in the " +
          "gray zone that needs prior coordination with AWS Trust & Safety.",
        topics: [
          'Shared responsibility revisited from an attacker\'s lens',
          'Control plane vs. data plane attacks',
          'Blast radius from service to account to organization',
          'The canonical AWS attack lifecycle',
          'Regulatory and contractual context for AWS engagements',
        ],
      },
      {
        n: 2,
        title: 'The AWS Attack Surface',
        pages: 14,
        description:
          "Enumerate the five entry classes of AWS compromise — public " +
          "misconfiguration, leaked credentials, identity-provider compromise, " +
          "compromised endpoint with cached credentials, and supply-chain into " +
          "AWS — and the STS trust primitives adversaries target. Introduces the " +
          "STS AssumeRole* family decision tree that drives Chapters 7 and 9, " +
          "covers IAM Identity Center as both target and control, walks the " +
          "regional vs. global service footprint that defines where an operator " +
          "must hunt, and catalogs the default configurations that bite " +
          "organizations whose engineers picked the AWS quickstart.",
        topics: [
          'The five entry classes',
          'STS and the AssumeRole* family decision tree',
          'IAM Identity Center (formerly AWS SSO)',
          'Regional vs. global service footprint',
          'Default configurations that bite',
        ],
      },
      {
        n: 3,
        title: 'Building an Offensive AWS Lab',
        pages: 14,
        description:
          "Stand up a multi-account AWS lab where the techniques in the rest " +
          "of the book can be detonated without risk to real data or budgets. " +
          "Covers the Organizations layout, Service Control Policies as " +
          "blast-radius guardrails, cost controls, CloudTrail on the attacker " +
          "account, the publicly available vulnerable-by-design Terraform " +
          "projects (CloudGoat, IAM Vulnerable) the rest of the book assumes are " +
          "deployed, a Terraform harness for rapid teardown, the logging and " +
          "visibility setup on the victim side, and a minimum operator " +
          "workstation build.",
        topics: [
          'Lab topology: management, victim, attacker OUs',
          'Service Control Policies as blast-radius guardrails',
          'Cost guardrails',
          'CloudTrail on the attacker account',
          'Vulnerable-by-design targets (CloudGoat, IAM Vulnerable)',
          'Terraform harness for rapid teardown',
          'Logging and visibility on the victim side',
          'A minimum operator workstation build',
        ],
      },
    ],
  },
  {
    section: 'Part II — Reconnaissance and Initial Access',
    chapters: [
      {
        n: 4,
        title: 'Unauthenticated Recon',
        pages: 18,
        description:
          "Methodically map an AWS footprint without any valid credentials. " +
          "Covers techniques that attribute an account ID from public artifacts " +
          "(S3 bucket names, AMI IDs, CloudFront distributions, SAML provider " +
          "certificates), enumerate principals by probing error responses, " +
          "brute-force role names through AssumeRole, identify exposed services " +
          "through public registries, chart DNS and CDN infrastructure, and " +
          "enumerate Cognito user pools. Closes with a one-hour external attack " +
          "surface workflow.",
        topics: [
          'Attribution — getting from domain to account',
          'Account-ID enumeration via quiet-riot and similar',
          'Role-name brute forcing through AssumeRole',
          'Public S3 discovery',
          'Public AMI, snapshot, and RDS enumeration',
          'DNS and CDN footprinting',
          'Lambda function URL and API Gateway discovery',
          'Cognito user-pool enumeration',
        ],
      },
      {
        n: 5,
        title: 'Authenticated Enumeration',
        pages: 18,
        description:
          "You have credentials — now translate them into a complete picture of " +
          "the account before attacking. Walks the essential aws-cli enumeration " +
          "discipline (the first five commands), CloudFox end-to-end as the " +
          "standard toolchain output, Pacu's enumeration modules for active " +
          "reconnaissance, and how to combine ScoutSuite, Prowler, and " +
          "CloudSploit findings into a prioritized attack-path list. Covers " +
          "PMapper privilege-escalation graphs, role trust policy reading, and " +
          "session-tag reconnaissance — turning a fresh credential pair into " +
          "an actionable attack plan while minimizing CloudTrail noise.",
        topics: [
          'The first five commands',
          'CloudFox end-to-end',
          'Pacu for active enumeration and exploit',
          'ScoutSuite, Prowler, and the compliance-adjacent tools',
          'PMapper for privilege-escalation graphs',
          'Reading role trust policies',
          'Session-tag reconnaissance',
        ],
      },
      {
        n: 6,
        title: 'Initial Access Vectors',
        pages: 16,
        description:
          "Walk through the six initial-access classes most productive in 2026: " +
          "leaked access keys in source and artifacts, keys baked into Docker " +
          "images and mobile applications, SSRF into IMDS (with the IMDSv2 " +
          "bypass landscape), phishing of federated console users, OAuth " +
          "application compromise, stolen developer-workstation credential " +
          "caches, and device-code phishing against IAM Identity Center. For " +
          "each: detection patterns, tooling, and the pivot path into an " +
          "authenticated session.",
        topics: [
          'Leaked access keys in source and artifacts',
          'Keys in Docker images',
          'Keys in mobile applications',
          'SSRF into IMDS (IMDSv1 and v2 bypass classes)',
          'Phishing federated console users',
          'Developer workstation credential caches',
          'Third-party OAuth application compromise',
          'Initial access recap and picking the right door',
        ],
      },
    ],
  },
  {
    section: 'Part III — IAM Exploitation',
    chapters: [
      {
        n: 7,
        title: 'IAM Deep Dive for Attackers',
        pages: 20,
        description:
          "Build an operator-grade mental model of how AWS evaluates an API " +
          "call. Covers the six policy types AWS consults (identity, resource, " +
          "organization, permission boundary, session, ACL), the explicit-deny " +
          "precedence, the canonical evaluation flowchart, principal types in " +
          "policy, the condition-key grammar that makes IAM both powerful and " +
          "routinely mis-written, the resource-policy surface, and the " +
          "role-trust primitives — ExternalId, SourceIdentity, the " +
          "confused-deputy problem, and Access Analyzer as defender radar — " +
          "that underpin Chapter 9. The flowchart in this chapter is the most " +
          "important diagram in the book.",
        topics: [
          'The six policy types',
          'The canonical evaluation flowchart',
          'Explicit deny precedence',
          'Principal types in policy',
          'Condition keys — the grammar that gets misused',
          'Resource policy surface',
          'Role trust policies revisited',
          'Confused-deputy problem and ExternalId',
          'SourceIdentity and the audit-trail angle',
          'Access Analyzer — defender radar you can read',
        ],
      },
      {
        n: 8,
        title: 'Privilege Escalation Techniques',
        pages: 20,
        description:
          "Walk through thirty-one canonical IAM privilege-escalation " +
          "primitives — the original Rhino Security Labs catalogue expanded " +
          "with current 2026 techniques. For each: prerequisites, exact API " +
          "sequence, detection opportunities, and the CloudGoat or IAM " +
          "Vulnerable scenario that demonstrates it. Closes with the " +
          "privilege-escalation matrix operators reach for in every " +
          "engagement, plus Pacu's iam__privesc_scan automation and the " +
          "detection strategy for blue teams.",
        topics: [
          'CreateAccessKey on another user; PassRole + compute',
          'AttachUserPolicy / AttachRolePolicy / AttachGroupPolicy',
          'PutUserPolicy / PutRolePolicy / PutGroupPolicy',
          'CreatePolicyVersion + SetAsDefault; SetDefaultPolicyVersion',
          'UpdateAssumeRolePolicy; UpdateLoginProfile; CreateLoginProfile',
          'AddUserToGroup; AssumeRole into over-permissive role',
          'Lambda UpdateFunctionCode and AddPermission',
          'CodeBuild UpdateProject + StartBuild; SSM SendCommand',
          'Glue UpdateDevEndpoint; CloudFormation drift takeover',
          'EC2 user-data overwrite; Parameter Store / Secrets replacement',
          'EventBridge rule hijack; S3 bucket-notification to attacker Lambda',
          'Step Functions state-machine rewrite',
          'IAM Identity Center permission-set escalation',
          'KMS grant to external principal; VPC endpoint policy relaxation',
          'Organizations SCP removal; service-linked role abuse',
          'Pacu iam__privesc_scan; the full primitive matrix',
          'Detection strategy for IAM privesc',
        ],
      },
      {
        n: 9,
        title: 'Cross-Account Attacks and Role Chaining',
        pages: 18,
        description:
          "Move from inside-one-account privesc to cross-account lateral " +
          "movement. Covers role chaining mechanics, the OrganizationAccount" +
          "AccessRole pathway that compromises an entire AWS Organization from " +
          "a single landing, cross-account sts:AssumeRole, resource-based-" +
          "policy cross-account abuse, AWS RAM resource shares, the confused-" +
          "deputy landscape before and after the March 2022 AWS behavior " +
          "change, resource-share footprinting, VPC peering as a network-plane " +
          "lateral vector, chained-session session-tag audit evasion, and the " +
          "full cross-account playbook stitching it all together.",
        topics: [
          'Role chaining mechanics',
          'The OrganizationAccountAccessRole pathway',
          'Cross-account sts:AssumeRole',
          'Resource-based policy cross-account abuse',
          'AWS RAM — shared resources across accounts',
          'The confused-deputy landscape after 2022',
          'Resource-share footprinting',
          'VPC peering and network plane',
          'Chained-session session tags and audit evasion',
          'The full cross-account playbook',
        ],
      },
    ],
  },
  {
    section: 'Part IV — Service-Specific Attacks',
    chapters: [
      {
        n: 10,
        title: 'Compute — EC2, SSM, and Lambda',
        pages: 20,
        description:
          "Compute is where instance metadata lives, where Session Manager " +
          "sessions originate, and where code runs on behalf of identities. " +
          "This chapter treats EC2, SSM, and Lambda together because each " +
          "attaches an IAM role to running code, and whoever controls the " +
          "code controls the role. Walks IMDS and the IMDSv2 bypass landscape, " +
          "the SSM agent as a pivot, SendCommand as remote code execution as " +
          "a service, the Lambda execution-role gold mine, Lambda layers and " +
          "dependency-confusion, and the CloudTrail/GuardDuty signatures " +
          "compute attacks leave behind.",
        topics: [
          'EC2 instance profiles and IMDS',
          'IMDSv2 bypass landscape',
          'The SSM agent as a pivot',
          'SendCommand: remote code execution as a service',
          'Lambda: the execution role gold mine',
          'Lambda layers and the dependency-confusion angle',
          'Compute detections and what evades them',
        ],
      },
      {
        n: 11,
        title: 'Storage — S3, EBS, and EFS',
        pages: 16,
        description:
          "The storage plane is where data lives, which makes it the highest-" +
          "leverage target after IAM itself. Covers the offensive vocabulary " +
          "for the three primary storage services: S3 object-access primitives " +
          "and resource-policy patterns that produce world-readable buckets; " +
          "EBS snapshot exfiltration as the favorite quiet-exfil pathway; the " +
          "EBS direct API for surgical block-level reads; the EFS NFS trust " +
          "surface and cross-account mount tradecraft. Each section closes " +
          "with the CloudTrail signature the technique leaves and the " +
          "GuardDuty findings that catch it.",
        topics: [
          'S3 object-access primitives',
          'EBS snapshot exfiltration',
          'EBS direct API for surgical reads',
          'EFS and the NFS trust surface',
          'Storage-plane detections',
        ],
      },
      {
        n: 12,
        title: 'Data Services — RDS, DynamoDB, Redshift, and Athena',
        pages: 16,
        description:
          "Managed data services are the primary targets in any data-centric " +
          "engagement. RDS holds the relational databases; DynamoDB holds " +
          "operational NoSQL; Redshift holds warehoused analytics; Athena " +
          "queries everything stored in S3. Each service has a distinct " +
          "exfiltration primitive that lets an authorized identity move data " +
          "without crossing the network perimeter at all: RDS snapshot " +
          "sharing, Aurora cross-account clones, DynamoDB export to S3, " +
          "Redshift UNLOAD to S3, and Athena CTAS. Walks each primitive " +
          "end-to-end with the IAM action set, the API sequence, the " +
          "CloudTrail signature, and post-exfil cleanup.",
        topics: [
          'RDS snapshot sharing',
          'Aurora and the cross-account clone',
          'DynamoDB export to S3',
          'Redshift UNLOAD',
          'Athena and the CTAS primitive',
          'RDS IAM auth and the password-less pivot',
          'Data-service detections',
        ],
      },
      {
        n: 13,
        title: 'Containers and Kubernetes — ECS, EKS, and Fargate',
        pages: 16,
        description:
          "Container workloads on AWS expose two identity surfaces: the " +
          "container's own task or pod identity, and the underlying compute " +
          "host's identity. This chapter walks both. For ECS: task roles, " +
          "task metadata, and the node-profile escape that turns a " +
          "compromised container on EC2 into a host-role pivot. For EKS: " +
          "IRSA (IAM Roles for Service Accounts), Pod Identity, the aws-auth " +
          "ConfigMap and its cluster-admin pivot, and the container escape " +
          "primitives that work in 2026. For ECR: image-pull poisoning. " +
          "Closes with the container-plane detections and what they miss.",
        topics: [
          'ECS task roles and task metadata',
          'ECS on EC2 and the node-profile escape',
          'EKS: IRSA and Pod Identity',
          'aws-auth and the cluster-admin pivot',
          'Container escape primitives',
          'ECR and image-pull poisoning',
          'Container-plane detections',
        ],
      },
      {
        n: 14,
        title: 'Serverless and Event-Driven',
        pages: 16,
        description:
          "Serverless and event-driven services are the connective tissue of " +
          "modern AWS architectures, and the trust relationships among them " +
          "are dense and often poorly understood. Walks the offensive surface " +
          "of API Gateway (authorization and the IAM proxy), EventBridge (the " +
          "persistence-friendly event router), SNS and SQS (resource-policy " +
          "abuse on queues and topics), Step Functions (state-machine pivot), " +
          "and the S3-notification primitive that turns a single object PUT " +
          "into attacker-controlled code execution. Emphasizes the event-" +
          "pattern subversion techniques that the defender's monitoring " +
          "frequently fails to catch.",
        topics: [
          'API Gateway: authorization and the IAM proxy',
          'EventBridge: the attack that looks like plumbing',
          'SNS and SQS resource-policy surface',
          'Step Functions and the state-machine pivot',
          'The S3-notification persistence primitive',
          'Event-pattern subversion: what the defender misses',
          'Serverless-plane detections',
        ],
      },
    ],
  },
  {
    section: 'Part V — Post-Exploitation',
    chapters: [
      {
        n: 15,
        title: 'Persistence',
        pages: 16,
        description:
          "MITRE ATT&CK catalogues nearly twenty cloud persistence techniques; " +
          "in AWS most collapse to one of three families. Identity persistence: " +
          "create or modify a principal — IAM user, access key, role, " +
          "federation trust — the attacker can use again later. Compute " +
          "persistence: create or modify an execution context — Lambda, " +
          "EventBridge rule, EC2 user-data, launch template — that runs " +
          "attacker code on a schedule or in response to an event. " +
          "Configuration persistence: create or modify a global account " +
          "setting — Identity Center permission sets, Organizations SCP, " +
          "service-level trusts — that broadens access in a way routine " +
          "credential rotation does not close.",
        topics: [
          'Identity persistence: users, keys, and MFA',
          'Federation-trust abuse: OIDC and SAML',
          'Compute persistence: Lambda, EventBridge, Launch Templates',
          'Configuration persistence: Identity Center, Organizations, SCPs',
          'A persistence matrix',
        ],
      },
      {
        n: 16,
        title: 'Defense Evasion',
        pages: 16,
        description:
          "Defense evasion in AWS is rarely about hiding processes or files " +
          "because the operator does not own the host on which their actions " +
          "execute — they are submitting authenticated API calls to AWS-" +
          "managed services. Evasion is therefore a CloudTrail and GuardDuty " +
          "discipline. Covers blending with automation, choosing CloudTrail-" +
          "friendly identities, controlling user-agent and source-IP " +
          "signatures, using session names and SourceIdentity that look " +
          "legitimate, regional arbitrage and log routing, narrowing the " +
          "audit log with unlogged actions, the cost of disabling or " +
          "diverting CloudTrail, GuardDuty specifics, and the surprising " +
          "counter-evasion risk that Access Analyzer poses to the attacker.",
        topics: [
          'Blending with automation',
          'User-agent and source-IP controls',
          'Session names and SourceIdentity',
          'Regional arbitrage and log routing',
          'Narrowing the log with unlogged actions',
          'Disabling and diverting CloudTrail',
          'GuardDuty specifics',
          'Access Analyzer as a counter-evasion risk',
        ],
      },
      {
        n: 17,
        title: 'Data Exfiltration and Impact',
        pages: 16,
        description:
          "Exfiltration in AWS is rarely about network tunnels because the " +
          "data plane is already service-to-service. Walks the categories " +
          "that work on a modern engagement: the secrets and KMS data-key " +
          "pattern; DNS exfiltration from inside AWS; the large-bulk exfil " +
          "playbook using cross-account snapshot sharing and DynamoDB export; " +
          "impact primitives (authorized only) — the destructive actions a " +
          "red team would document but not execute. Every category is " +
          "annotated with the evidence trail it leaves and the MITRE ATT&CK " +
          "technique it maps to.",
        topics: [
          'Exfiltration categories',
          'Secrets and the KMS data-key pattern',
          'DNS exfiltration from inside AWS',
          'The large-bulk exfil playbook',
          'Impact primitives (authorized only)',
          'Exfiltration evidence trail',
          'MITRE ATT&CK alignment',
        ],
      },
    ],
  },
  {
    section: 'Part VI — Tradecraft and Reporting',
    chapters: [
      {
        n: 18,
        title: 'The Operator\'s Toolbox',
        pages: 18,
        description:
          "A mature AWS engagement uses a small number of tools intensively " +
          "rather than a large number of tools superficially. This chapter " +
          "inventories the tools referenced throughout the book, plus a " +
          "handful not yet introduced, and describes when each one earns a " +
          "place on the engagement laptop. Categorizes by phase " +
          "(enumeration / reconnaissance, secret hunting / credential " +
          "discovery, exploitation / post-exploitation), ends with a curated " +
          "tool matrix and a build for an operator's own harness. Lineup " +
          "evolves year by year; this reflects the 2026 landscape.",
        topics: [
          'Enumeration and reconnaissance tools',
          'Secret hunting and credential discovery',
          'Exploitation and post-exploitation tools',
          'A curated tool matrix',
          'Building an operator\'s harness',
        ],
      },
      {
        n: 19,
        title: 'CI/CD and Supply-Chain Attacks into AWS',
        pages: 16,
        description:
          "CI/CD pipelines are now one of the most reliable entry points into " +
          "AWS. The reason is structural: most pipelines hold IAM credentials " +
          "of some form (long-lived keys in older setups, OIDC-issued " +
          "short-term sessions in modern ones), they execute attacker-" +
          "controllable code (build steps), and they are connected to " +
          "production deploy paths. Walks the dominant 2026 supply-chain " +
          "surface: GitHub Actions OIDC trust abuse, Terraform state as an " +
          "attacker resource, dependency confusion against Lambda builds, " +
          "GitHub Actions workflow injection, the CodeBuild + IAM PassRole " +
          "bridge, and IaC pipeline injection writ large.",
        topics: [
          'The GitHub Actions OIDC trust',
          'Terraform state as an attacker resource',
          'Dependency confusion against Lambda builds',
          'GitHub Actions workflow injection',
          'CodeBuild, CodePipeline, and the IAM-PassRole bridge',
          'IaC pipeline injection writ large',
        ],
      },
      {
        n: 20,
        title: 'Reporting, Detection, and Remediation',
        pages: 18,
        description:
          "Findings without context are invoices; findings with context are " +
          "leverage. Walks the practitioner deliverables that determine " +
          "whether an engagement produces lasting value: the report " +
          "structure expected by mature security organizations, the anatomy " +
          "of an individual finding (what, why, evidence, severity, " +
          "remediation), the CloudTrail evidence package per finding, the " +
          "detection-engineering deliverable (queries and SIEM rules), " +
          "Terraform-as-remediation, severity calibration, and the defender " +
          "hand-off. Closes the loop on every prior chapter — every primitive " +
          "in chapters 4-19 plugs into a finding template.",
        topics: [
          'Report structure',
          'Finding anatomy',
          'CloudTrail evidence per finding',
          'Detection engineering deliverables',
          'Remediation as Terraform',
          'Severity calibration',
          'Defender hand-off',
        ],
      },
    ],
  },
];

// Appendices and back matter (rolled into the last "section")
const APPENDICES = [
  { code: 'A', title: 'Command Cheat Sheets by Service',           pages: 14 },
  { code: 'B', title: 'IAM Privilege-Escalation Matrix',           pages: 8 },
  { code: 'C', title: 'MITRE ATT&CK for Cloud — Mapping',          pages: 10 },
  { code: 'D', title: 'CloudTrail + Athena Detection Queries',     pages: 10 },
  { code: '-', title: 'Glossary',                                  pages: 8 },
  { code: '-', title: 'Bibliography and References',               pages: 6 },
  { code: '-', title: 'Index',                                     pages: 8 },
];

// ----------------------------------------------------------------------------
// Build the document body
// ----------------------------------------------------------------------------
const body = [];

// ---------- HEADER / SECTION I ---------------------------------------------
body.push(pageTitle('Outline Template'));
body.push(sectionTitle('Section I'));

// Book Title
body.push(inlineLabelValue('Book Title', 'AWS Offensive Security'));
body.push(italicHint('Mention the technology and the concept the book is based on.'));

// Subtitle
body.push(inlineLabelValue('Subtitle', subtitle));
body.push(italicHint(
  'This should be an extension of the title of the book. Something that can ' +
  'create a spark in the audience. Try using SEO friendly keywords.'
));
body.push(smallNote('Character count (with spaces): ' + subtitle.length + ' / 70'));

// Who is this book for
body.push(inlineLabelValue('Who is this book for', audience));
body.push(italicHint(
  'Audience profile: experience level, professional roles, prerequisites.'
));
body.push(smallNote('Character count (with spaces): ' + audience.length + ' / 350'));

// Book Description
body.push(pinkHeading('Book Description'));
description.forEach(p => body.push(P(p, { spacing: { after: 160 } })));
const descLen = description.join(' ').length;
body.push(smallNote('Character count (with spaces): ' + descLen + ' / 1000-1350'));

// Key Features
body.push(pinkHeading('Key Features'));
body.push(italicHint('(Five points; ≤70 characters each.)'));
keyFeatures.forEach(k => body.push(bullet(k)));

// Competitor Analysis
body.push(pinkHeading('Competitor Analysis'));
body.push(P(
  'The two most directly comparable titles in print today are listed below, ' +
  'with a positioning note for each.'
));
competitorAnalysis.forEach((c, i) => {
  body.push(new Paragraph({
    spacing: { before: 160, after: 60 },
    children: [
      new TextRun({ text: `Competing book ${i + 1}: `, bold: true, size: SIZE_BODY, color: BPB_PINK, font: FONT }),
      new TextRun({ text: c.title, bold: true,        size: SIZE_BODY, color: BODY_INK, font: FONT }),
      new TextRun({ text: ` — ${c.author}`, italics: true, size: SIZE_BODY, color: BODY_INK, font: FONT }),
    ],
  }));
  body.push(P(c.note));
});
body.push(P(competitorPositioning, { spacing: { before: 120 } }));

// Tech List
body.push(pinkHeading('Tech List'));
body.push(italicHint(
  'Libraries / tools / software / platforms / programming languages used in the book.'
));
techList.forEach(t => body.push(P(t, { spacing: { after: 120 } })));

// Author Bio
body.push(pinkHeading('Author Bio'));
authorBio.forEach(p => body.push(P(p, { spacing: { after: 160 } })));

// ---------- SECTION II: Table of Contents ----------------------------------
body.push(new Paragraph({ children: [new PageBreak()] }));
body.push(sectionTitle('Section II'));
body.push(pinkHeading('Table of Contents'));
body.push(italicHint(
  'Section / Chapter / Page Count / Delivery Date — page counts are estimates ' +
  'based on the manuscript draft; delivery dates to be filled in once the ' +
  'contract schedule is finalized.'
));

// Build the TOC table
const colWidths = [2200, 800, 4760, 800, 800];   // sums to 9360 (US Letter content width)
const tocRows = [];

// Header row
tocRows.push(new TableRow({
  tableHeader: true,
  children: [
    tableCell('Section Title',     { width: colWidths[0], bold: true, fill: TABLE_HEADER_BG, center: true }),
    tableCell('Chapter No',        { width: colWidths[1], bold: true, fill: TABLE_HEADER_BG, center: true }),
    tableCell('Chapter Title',     { width: colWidths[2], bold: true, fill: TABLE_HEADER_BG, center: true }),
    tableCell('Page Count',        { width: colWidths[3], bold: true, fill: TABLE_HEADER_BG, center: true }),
    tableCell('Delivery Date',     { width: colWidths[4], bold: true, fill: TABLE_HEADER_BG, center: true }),
  ],
}));

let totalChapterPages = 0;
PARTS.forEach(part => {
  part.chapters.forEach((ch, idx) => {
    totalChapterPages += ch.pages;
    tocRows.push(new TableRow({
      children: [
        // Only put the section title in the first row of each section
        tableCell(idx === 0 ? part.section : '', { width: colWidths[0], bold: idx === 0 }),
        tableCell(String(ch.n),                  { width: colWidths[1], center: true }),
        tableCell(ch.title,                      { width: colWidths[2] }),
        tableCell(String(ch.pages),              { width: colWidths[3], center: true }),
        tableCell('TBD',                         { width: colWidths[4], center: true, italics: true, color: ITALIC_HINT }),
      ],
    }));
  });
});

// Appendices block
APPENDICES.forEach((a, idx) => {
  totalChapterPages += a.pages;
  tocRows.push(new TableRow({
    children: [
      tableCell(idx === 0 ? 'Appendices & Back Matter' : '', { width: colWidths[0], bold: idx === 0 }),
      tableCell(a.code,        { width: colWidths[1], center: true }),
      tableCell(a.title,       { width: colWidths[2] }),
      tableCell(String(a.pages), { width: colWidths[3], center: true }),
      tableCell('TBD',         { width: colWidths[4], center: true, italics: true, color: ITALIC_HINT }),
    ],
  }));
});

// Total row
tocRows.push(new TableRow({
  children: [
    tableCell('', { width: colWidths[0] }),
    tableCell('', { width: colWidths[1] }),
    tableCell('Total Pages (estimated)', { width: colWidths[2], bold: true, italics: true, fill: TABLE_HEADER_BG }),
    tableCell(String(totalChapterPages), { width: colWidths[3], bold: true, center: true, fill: TABLE_HEADER_BG }),
    tableCell('', { width: colWidths[4], fill: TABLE_HEADER_BG }),
  ],
}));

body.push(new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: colWidths,
  rows: tocRows,
}));

// ---------- CHAPTER DETAILS ------------------------------------------------
body.push(new Paragraph({ children: [new PageBreak()] }));
body.push(pinkHeading('Chapter Details'));
body.push(italicHint(
  'Each chapter description is 100-200 words and lists the topics covered. ' +
  'Counts and topics are derived from the current manuscript draft.'
));

PARTS.forEach(part => {
  // Section divider — same styling as standalone pink heading (14pt bold pink)
  body.push(new Paragraph({
    spacing: { before: 280, after: 140 },
    children: [new TextRun({
      text: part.section,
      bold: true, size: SIZE_PINK_HEAD, color: BPB_PINK, font: FONT,
    })],
  }));

  part.chapters.forEach(ch => {
    // Chapter heading — matches the template's "Chapter X: [Chapter Title]" style: 14pt bold pink
    body.push(new Paragraph({
      spacing: { before: 220, after: 80 },
      children: [new TextRun({
        text: `Chapter ${ch.n}: ${ch.title} [${ch.pages} pages]`,
        bold: true, size: SIZE_PINK_HEAD, color: BPB_PINK, font: FONT,
      })],
    }));

    // Description — body text, label bold (same family as the template's "Description:" line)
    body.push(new Paragraph({
      spacing: { after: 120 },
      children: [
        new TextRun({ text: 'Description: ', bold: true, size: SIZE_BODY, color: BODY_INK, font: FONT }),
        new TextRun({ text: ch.description,  size: SIZE_BODY, color: BODY_INK, font: FONT }),
      ],
    }));
    const wc = ch.description.split(/\s+/).filter(Boolean).length;
    body.push(smallNote('Word count: ' + wc + ' / 100-200'));

    // Topics — body text, label bold (matches the template)
    body.push(new Paragraph({
      spacing: { before: 80, after: 60 },
      children: [
        new TextRun({ text: 'Topics to be covered:', bold: true, size: SIZE_BODY, color: BODY_INK, font: FONT }),
      ],
    }));
    ch.topics.forEach(t => body.push(bullet(t)));
  });
});

// ----------------------------------------------------------------------------
// Document
// ----------------------------------------------------------------------------
const doc = new Document({
  creator: 'AWS Offensive Security book project',
  title: 'BPB Book Outline — AWS Offensive Security',
  description: 'Filled-out BPB Publications book outline template.',
  styles: {
    default: { document: { run: { font: FONT, size: SIZE_BODY } } },
  },
  numbering: {
    config: [{
      reference: 'bpb-bullets',
      levels: [{
        level: 0,
        format: LevelFormat.BULLET,
        text: '➢',                    // ➢ — matches the BPB template's arrow-bullet style
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } },
      }],
    }],
  },
  sections: [{
    properties: {
      page: {
        size: { width: PAGE_W, height: PAGE_H },
        margin: { top: MARG, right: MARG, bottom: MARG, left: MARG },
      },
    },
    headers: {
      default: new Header({ children: [new Paragraph({
        alignment: AlignmentType.RIGHT,
        children: [new TextRun({
          text: 'BPB Book Outline — AWS Offensive Security',
          italics: true, size: 18, color: ITALIC_HINT, font: FONT,
        })],
      })] }),
    },
    footers: {
      default: new Footer({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({ text: 'Page ', size: 18, color: ITALIC_HINT, font: FONT }),
          new TextRun({ children: [PageNumber.CURRENT], size: 18, color: ITALIC_HINT, font: FONT }),
          new TextRun({ text: ' of ',  size: 18, color: ITALIC_HINT, font: FONT }),
          new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, color: ITALIC_HINT, font: FONT }),
        ],
      })] }),
    },
    children: body,
  }],
});

// Write the file next to this script
const outPath = path.join(__dirname, 'BPB_Book_Outline_FILLED.docx');
Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(outPath, buf);
  console.log('[+] Wrote ' + outPath);
  console.log('    Total chapter pages (estimated): ' + totalChapterPages);
}).catch(err => {
  console.error('[-] Failed to build outline:', err);
  process.exit(1);
});
